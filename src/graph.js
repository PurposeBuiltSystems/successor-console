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
      return (await withTimeout(pca.acquireTokenSilent({ scopes: SCOPES }), 20000, "silent timeout")).accessToken;
    } catch (e) {
      var interactive = await withTimeout(
        pca.acquireTokenPopup({ scopes: SCOPES }), 120000,
        "Sign-in didn't finish. A Microsoft sign-in window may have opened behind Outlook — " +
        "check for it (or Mission Control), finish signing in, and click again. If no window " +
        "appeared at all, fully quit Outlook (Cmd+Q), reopen, and retry.");
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
