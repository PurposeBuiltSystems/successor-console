/*
 * Successor Console — Microsoft Graph data layer.
 *
 * AUTH: Nested App Authentication (NAA) via MSAL — no backend; identical
 * pattern to the other PurposeBuilt add-ins. Everything is DELEGATED:
 * the inherited mailbox is read via Mail.Read.Shared, which only works for
 * mailboxes the signed-in user has ALREADY been granted access to by IT
 * (typically a departed employee's mailbox converted to a shared mailbox
 * with successor Full Access). The add-in can never reach a mailbox the
 * user couldn't open in Outlook themselves.
 *
 * Exposes a global `GraphData` object.
 */
/* global msal */
(function (root) {
  "use strict";

  var CLIENT_ID = "74663a5f-b6db-4c06-83a0-325fb26d7a55"; // "Successor Console" Entra app (purposebuilt.systems tenant)
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var SCOPES = [
    "Mail.Read",          // own mailbox (when auditing your own archive)
    "Mail.Read.Shared",   // the inherited/shared mailbox
    "Files.ReadWrite.All",// write the pack to OneDrive or a SharePoint folder
  ];

  var pcaPromise = null;

  function getPca() {
    if (!pcaPromise) {
      pcaPromise = msal.createNestablePublicClientApplication({
        auth: {
          clientId: CLIENT_ID,
          authority: "https://login.microsoftonline.com/common",
        },
      });
    }
    return pcaPromise;
  }

  // --- add-in sign-out state -------------------------------------------
  //
  // Certification rejected the naive version on a sibling add-in: "after
  // clicking sign-out there is no response or not signed out." The reason is
  // structural. Under nested app authentication Outlook owns the session and
  // getAllAccounts() reports the HOST's account, not a cache this add-in
  // controls - so clearing MSAL's cache changes nothing visible, the next
  // silent acquisition succeeds anyway, and the pane redraws as signed in.
  //
  // A sign-out this add-in cannot deliver should not be offered. What it CAN
  // deliver is refusing to act until the user authenticates again: while
  // signed out it reports itself signed out and will not use a silent token,
  // so the next action raises a real prompt. Outlook's own session is
  // untouched, and the pane says so.
  var SIGNED_OUT_KEY = "addinSignedOut";
  var signedOut = false;
  try { signedOut = Office.context.roamingSettings.get(SIGNED_OUT_KEY) === true; } catch (e) { signedOut = false; }

  function setSignedOut(v) {
    signedOut = !!v;
    try {
      Office.context.roamingSettings.set(SIGNED_OUT_KEY, signedOut);
      Office.context.roamingSettings.saveAsync(function () {});
    } catch (e) { /* in-memory is still correct for this session */ }
  }

  /** Signed-in account, or null. Reports null while signed out, by design. */
  async function currentAccount() {
    if (signedOut) { return null; }
    try {
      var pca = await getPca();
      var accts = (pca.getAllAccounts && pca.getAllAccounts()) || [];
      return accts.length ? (accts[0].username || accts[0].name || "signed in") : null;
    } catch (e) { return null; }
  }

  /**
   * Sign out of the add-in. The state flips SYNCHRONOUSLY before any awaiting
   * so the pane can respond instantly - awaiting a broker handshake first is
   * the "no response" half of the finding. Cache clearing is best-effort on
   * top; the enforced state is what makes this real.
   */
  function signOut() {
    setSignedOut(true);
    var pending = pcaPromise;      // only clear what exists; never start a handshake here
    pcaPromise = null;
    if (!pending) { return Promise.resolve(true); }
    return Promise.resolve(pending).then(function (pca) {
      var accts = (pca && pca.getAllAccounts && pca.getAllAccounts()) || [];
      var chain = Promise.resolve();
      accts.forEach(function (a) {
        chain = chain.then(function () {
          if (pca.clearCache) { return pca.clearCache({ account: a }); }
          if (pca.logoutPopup) { return pca.logoutPopup({ account: a }); }
        }).catch(function () { /* best effort; the enforced state stands */ });
      });
      return chain.then(function () { return true; });
    }).catch(function () { return true; });
  }

  /** True while the user has signed the add-in out. */
  function isSignedOut() { return signedOut; }




  /**
   * Sign-in must never hang the pane. An un-timed await on the popup flow
   * leaves a button disabled with nothing visible happening — which reads
   * to the user as "the button does nothing" and gives them nothing to act
   * on. Fail loudly instead, naming the two things that actually fix it.
   */
  function withTimeout(promise, ms, message) {
    var timer;
    return Promise.race([
      promise.then(function (v) { clearTimeout(timer); return v; },
                   function (e) { clearTimeout(timer); throw e; }),
      new Promise(function (_, reject) {
        timer = setTimeout(function () { reject(new Error(message)); }, ms);
      }),
    ]);
  }

  async function getToken() {
    var pca = await withTimeout(getPca(), 20000,
      "Sign-in didn't start. Fully quit Outlook (Cmd+Q) and reopen, then try again.");
    try {
      // Signed out means signed out: skip silent so the user must re-authenticate.
      if (signedOut) { throw new Error("signed out of the add-in"); }
      return (await withTimeout(pca.acquireTokenSilent({ scopes: SCOPES }), 20000, "silent timeout")).accessToken;
    } catch (e) {
      var interactive = await withTimeout(
        pca.acquireTokenPopup({ scopes: SCOPES }), 120000,
        "Sign-in didn't finish. A Microsoft sign-in window may have opened behind Outlook — " +
        "check for it (or Mission Control), finish signing in, and click again. If no window " +
        "appeared at all, fully quit Outlook (Cmd+Q), reopen, and retry.");
      setSignedOut(false);   // a real interactive sign-in ends the signed-out state
      return interactive.accessToken;
    }
  }

  async function graph(token, method, path, body, raw) {
    var res = await fetch(GRAPH + path, {
      method: method,
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": raw ? "application/octet-stream" : "application/json",
      },
      body: raw ? body : (body ? JSON.stringify(body) : undefined),
    });
    if (!res.ok) {
      var text = await res.text();
      throw new Error("Graph " + method + " " + path + " -> " + res.status + " " + text);
    }
    return res.status === 204 ? null : res.json();
  }

  /** Page through a Graph collection following @odata.nextLink. */
  async function graphAll(token, path, cap) {
    var items = [];
    var url = GRAPH + path;
    var guard = 0;
    while (url && guard++ < (cap || 40)) {
      var res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
      if (!res.ok) { throw new Error("Graph GET " + url + " -> " + res.status); }
      var page = await res.json();
      items = items.concat(page.value || []);
      url = page["@odata.nextLink"] || null;
    }
    return items;
  }

  /* ------------------------------------------------------------- mailbox */

  /** "" or null = the signed-in user's own mailbox; otherwise a shared
   *  mailbox address the user has Full Access to (Mail.Read.Shared). */
  function mailboxBase(address) {
    var a = String(address || "").trim();
    return a ? "/users/" + encodeURIComponent(a) : "/me";
  }

  var MSG_SELECT = "id,conversationId,subject,bodyPreview,webLink";

  async function sentMessages(token, address, daysBack) {
    var since = new Date(Date.now() - daysBack * 864e5).toISOString();
    return graphAll(token,
      mailboxBase(address) + "/mailFolders/sentitems/messages" +
      "?$select=" + MSG_SELECT + ",sentDateTime,toRecipients,ccRecipients" +
      "&$filter=sentDateTime ge " + since + "&$top=100");
  }

  async function inboundMessages(token, address, daysBack) {
    var since = new Date(Date.now() - daysBack * 864e5).toISOString();
    return graphAll(token,
      mailboxBase(address) + "/messages" +
      "?$select=" + MSG_SELECT + ",receivedDateTime,from" +
      "&$filter=receivedDateTime ge " + since + "&$top=100");
  }

  /** Messages with attachments (metadata only — contentBytes never fetched). */
  async function attachmentMessages(token, address, daysBack) {
    var since = new Date(Date.now() - daysBack * 864e5).toISOString();
    return graphAll(token,
      mailboxBase(address) + "/messages" +
      "?$select=id,subject,receivedDateTime,from,webLink,hasAttachments" +
      "&$expand=attachments($select=id,name,size,contentType,isInline)" +
      "&$filter=receivedDateTime ge " + since + " and hasAttachments eq true&$top=50", 20);
  }

  /** Hydrate full bodies for chosen precedent answers (capped by caller). */
  async function messageBody(token, address, id) {
    var m = await graph(token, "GET",
      mailboxBase(address) + "/messages/" + id + "?$select=body");
    return (m && m.body && m.body.content) || "";
  }

  /* --------------------------------------------------------------- drive */

  /**
   * OneDrive preflight (Records Packager lesson): a licensed-but-never-
   * opened OneDrive 404s on /me/drive until the user visits office.com once.
   */
  async function ensureDrive(token) {
    try {
      return await graph(token, "GET", "/me/drive?$select=id,webUrl");
    } catch (e) {
      if (/404/.test(String(e && e.message))) {
        throw new Error(
          "Your OneDrive hasn't been set up yet. Open onedrive.com or " +
          "office.com once, open OneDrive there, then try again — or paste " +
          "a SharePoint folder link instead.");
      }
      throw e;
    }
  }

  /** Resolve a pasted SharePoint/OneDrive share URL to a writable folder. */
  async function resolveShareFolder(token, shareUrl) {
    var b64 = btoa(unescape(encodeURIComponent(shareUrl)))
      .replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
    var item = await graph(token, "GET",
      "/shares/u!" + b64 + "/driveItem?$select=id,name,webUrl,parentReference,folder,sharepointIds");
    if (!item.folder) { throw new Error("That link points to a file — paste a link to a FOLDER."); }
    return {
      driveId: item.parentReference.driveId,
      itemId: item.id,
      webUrl: item.webUrl,
      sharepointIds: item.sharepointIds || null,
    };
  }

  /** Explicit chain creation (conflictBehavior fail + swallow 409 —
   *  Records Packager lesson: don't rely on PUT-by-path auto-creating). */
  async function ensureFolder(token, driveId, parentId, name) {
    try {
      var made = await graph(token, "POST",
        "/drives/" + driveId + "/items/" + parentId + "/children",
        { name: name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" });
      return made.id;
    } catch (e) {
      if (!/409|nameAlreadyExists/.test(String(e && e.message))) { throw e; }
      var kids = await graphAll(token,
        "/drives/" + driveId + "/items/" + parentId + "/children?$select=id,name,folder&$top=200", 5);
      for (var i = 0; i < kids.length; i++) {
        if (kids[i].name === name && kids[i].folder) { return kids[i].id; }
      }
      throw e;
    }
  }

  /** Simple PUT upload (pack files are small; <4 MB path is fine). */
  async function uploadFile(token, driveId, parentId, name, content) {
    return graph(token, "PUT",
      "/drives/" + driveId + "/items/" + parentId + ":/" + encodeURIComponent(name) + ":/content",
      content, true);
  }

  /** sharepointIds for the pack folder — feeds the .agent items_by_url. */
  async function folderInfo(token, driveId, itemId) {
    return graph(token, "GET",
      "/drives/" + driveId + "/items/" + itemId + "?$select=id,webUrl,sharepointIds");
  }

  root.GraphData = {
    signOut: signOut,
    currentAccount: currentAccount,
    isSignedOut: isSignedOut,
    getToken: getToken,
    sentMessages: sentMessages,
    inboundMessages: inboundMessages,
    attachmentMessages: attachmentMessages,
    messageBody: messageBody,
    ensureDrive: ensureDrive,
    resolveShareFolder: resolveShareFolder,
    ensureFolder: ensureFolder,
    uploadFile: uploadFile,
    folderInfo: folderInfo,
    _config: { clientId: CLIENT_ID },
  };
})(typeof self !== "undefined" ? self : this);
