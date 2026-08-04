/*
 * Successor Console — task pane UI wiring.
 *
 * Scan: pull the inherited mailbox (sent + inbound + attachment metadata)
 * via GraphData, run the pure Derive analyses, browse the results in tabs.
 * Build: assemble the knowledge pack file plan (Pack), render .docx parts
 * (Docx + JSZip), upload everything to OneDrive or a pasted SharePoint
 * folder, optionally drop a best-effort "Ask the Archive.agent" alongside.
 */
/* global Office, GraphData, Derive, Docx, Pack, JSZip, document, window, TextEncoder */
(function () {
  "use strict";

  var scanData = null; // results of the last scan
  var activeTab = "people";

  Office.onReady(function () {
    on("scan", "click", scan);
    on("build", "click", build);
    on("search", "input", renderTab);
    Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (b) {
      b.addEventListener("click", function () {
        activeTab = b.getAttribute("data-tab");
        Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (x) {
          x.className = "tab" + (x === b ? " active" : "");
        });
        renderTab();
      });
    });
  });

  function byId(id) { return document.getElementById(id); }

  /**
   * Outlook caches the pane HTML but the ?v= query string makes it fetch
   * JavaScript fresh, so a returning user can run today's JS against
   * yesterday's page. Binding through this helper means a missing element
   * costs one feature instead of throwing and leaving every later button
   * unbound — a whole dead pane.
   */
  function on(id, ev, fn) {
    var el = byId(id);
    if (el) { el.addEventListener(ev, fn); }
    return el;
  }

  function setStatus(kind, text) {
    var el = byId("status");
    if (!text) { el.hidden = true; return; }
    el.hidden = false;
    el.className = "status " + kind;
    el.textContent = text;
  }

  /* ------------------------------------------------------------------ scan */

  async function scan() {
    var address = byId("mailbox").value.trim();
    var years = Math.max(1, Math.min(10, parseInt(byId("yearsBack").value, 10) || 3));
    var maxPairs = Math.max(20, Math.min(500, parseInt(byId("maxPairs").value, 10) || 200));
    var daysBack = years * 365;
    byId("scan").disabled = true;
    byId("results").hidden = true;
    try {
      var token = await GraphData.getToken();
      var me = ((Office.context.mailbox.userProfile || {}).emailAddress || "").toLowerCase();
      var owner = address || me;

      setStatus("work", "Reading sent mail from " + owner + "…");
      var sent = await GraphData.sentMessages(token, address, daysBack);
      setStatus("work", sent.length + " sent — reading received mail…");
      var inbound = await GraphData.inboundMessages(token, address, daysBack);
      setStatus("work", inbound.length + " received — indexing attachments…");
      var attachMsgs = await GraphData.attachmentMessages(token, address, daysBack);

      setStatus("work", "Deriving…");
      var correspondents = Derive.rankCorrespondents(sent, inbound, owner);
      var topics = Derive.clusterTopics(sent, inbound);
      var open = Derive.findOpenItems(sent, inbound, {});
      var precedents = Derive.extractPrecedents(sent, inbound, { maxPairs: maxPairs });

      var attachments = [];
      attachMsgs.forEach(function (m) {
        (m.attachments || []).forEach(function (a) {
          if (a.isInline) { return; }
          attachments.push({
            name: a.name || "(unnamed)", size: a.size || 0,
            date: m.receivedDateTime, subject: m.subject || "(no subject)",
            from: ((m.from || {}).emailAddress || {}).name || ((m.from || {}).emailAddress || {}).address || "",
            webLink: m.webLink || "",
          });
        });
      });
      attachments.sort(function (a, b) { return a.date < b.date ? 1 : -1; });

      scanData = {
        address: address, owner: owner, years: years,
        sent: sent, inbound: inbound,
        correspondents: correspondents, topics: topics, open: open,
        precedents: precedents, attachments: attachments,
      };
      if (!byId("ownerName").value.trim()) {
        byId("ownerName").value = owner.split("@")[0].replace(/[._]/g, " ")
          .replace(/\b\w/g, function (c) { return c.toUpperCase(); });
      }
      byId("results").hidden = false;
      renderTab();
      setStatus("info", correspondents.length + " correspondents · " + topics.length +
        " topics · " + (open.waiting.length + open.owed.length) + " open items · " +
        precedents.length + " precedent Q&A pairs · " + attachments.length +
        " attachments (" + years + " yr scan).");
    } catch (e) {
      var msg = (e && e.message) || String(e);
      if (/REPLACE_WITH_ENTRA_CLIENT_ID/.test(GraphData._config.clientId)) {
        msg = "Set your Entra client ID in src/graph.js before running. (" + msg + ")";
      } else if (/ErrorAccessDenied|403/.test(msg) && address) {
        msg = "Access denied to " + address + " — you need Full Access to that mailbox " +
          "(ask IT), and the address must be exact. (" + msg + ")";
      }
      setStatus("error", "Scan failed: " + msg);
    } finally {
      byId("scan").disabled = false;
    }
  }

  /* ---------------------------------------------------------------- browse */

  function entry(host, title, meta, detail, webLink) {
    var div = document.createElement("div");
    div.className = "entry";
    var top = document.createElement("div");
    top.className = "top";
    var t = document.createElement("div");
    t.className = "title";
    if (webLink) {
      var a = document.createElement("a");
      a.href = webLink;
      a.target = "_blank";
      a.rel = "noopener";
      a.textContent = title;
      t.appendChild(a);
    } else {
      t.textContent = title;
    }
    var m = document.createElement("div");
    m.className = "meta";
    m.textContent = meta || "";
    top.appendChild(t);
    top.appendChild(m);
    div.appendChild(top);
    if (detail) {
      var d = document.createElement("div");
      d.className = "detail";
      d.textContent = detail;
      div.appendChild(d);
    }
    host.appendChild(div);
  }

  function renderTab() {
    if (!scanData) { return; }
    var host = byId("tabContent");
    host.innerHTML = "";
    var q = byId("search").value.trim().toLowerCase();
    function match(s) { return !q || String(s).toLowerCase().indexOf(q) !== -1; }

    if (activeTab === "people") {
      scanData.correspondents.slice(0, 200).forEach(function (p) {
        if (!match(p.name + " " + p.email)) { return; }
        entry(host, p.name, p.total + " emails",
          p.email + " · " + p.first.slice(0, 10) + " → " + p.last.slice(0, 10));
      });
    } else if (activeTab === "open") {
      scanData.open.waiting.forEach(function (it) {
        if (!match(it.subject + " " + it.who)) { return; }
        entry(host, it.subject, "⏳ " + it.date.slice(0, 10),
          "They never answered — asked of " + it.who + ": " + it.text.slice(0, 160), it.webLink);
      });
      scanData.open.owed.forEach(function (it) {
        if (!match(it.subject + " " + it.who)) { return; }
        entry(host, it.subject, "📥 " + it.date.slice(0, 10),
          "Never answered — asked by " + it.who + ": " + it.text.slice(0, 160), it.webLink);
      });
    } else if (activeTab === "topics") {
      scanData.topics.slice(0, 200).forEach(function (t) {
        if (!match(t.label)) { return; }
        entry(host, t.label, t.count + " emails",
          t.conversations + " conversations · " + t.first.slice(0, 10) + " → " + t.last.slice(0, 10));
      });
    } else if (activeTab === "files") {
      scanData.attachments.slice(0, 300).forEach(function (a) {
        if (!match(a.name + " " + a.subject + " " + a.from)) { return; }
        entry(host, a.name, Math.round(a.size / 1024) + " KB",
          a.date.slice(0, 10) + " · from " + a.from + " · “" + a.subject + "”", a.webLink);
      });
    }
    if (!host.children.length) {
      var p = document.createElement("div");
      p.className = "detail";
      p.textContent = "Nothing matches.";
      host.appendChild(p);
    }
  }

  /* ----------------------------------------------------------------- build */

  async function docxBuffer(doc) {
    var built = Docx.buildDocx(doc);
    var zip = new JSZip();
    Object.keys(built.parts).forEach(function (path) { zip.file(path, built.parts[path]); });
    return zip.generateAsync({
      type: "arraybuffer",
      compression: "DEFLATE",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
  }

  async function build() {
    if (!scanData) { return; }
    var ownerName = byId("ownerName").value.trim() || scanData.owner.split("@")[0];
    var destUrl = byId("destUrl").value.trim();
    var wantAgent = byId("emitAgent").checked;
    byId("build").disabled = true;
    byId("packResult").hidden = true;
    try {
      var token = await GraphData.getToken();

      // destination root
      setStatus("work", "Resolving the destination…");
      var driveId, rootId, isSharePoint = false;
      if (destUrl) {
        var share = await GraphData.resolveShareFolder(token, destUrl);
        driveId = share.driveId;
        rootId = share.itemId;
        isSharePoint = !!(share.sharepointIds && share.sharepointIds.siteId);
      } else {
        var drive = await GraphData.ensureDrive(token);
        driveId = drive.id;
        var root = await GraphData.ensureFolder(token, driveId, "root", "Knowledge Packs");
        rootId = root;
      }
      var packFolderId = await GraphData.ensureFolder(token, driveId, rootId,
        "Knowledge Pack - " + ownerName);
      var precFolderId = await GraphData.ensureFolder(token, driveId, packFolderId, "precedents");

      // hydrate full bodies for the answers that made the cut (capped)
      var toHydrate = scanData.precedents.slice(0, 60);
      for (var h = 0; h < toHydrate.length; h++) {
        setStatus("work", "Fetching answer text " + (h + 1) + "/" + toHydrate.length + "…");
        try {
          var html = await GraphData.messageBody(token, scanData.address, toHydrate[h].answer.id);
          toHydrate[h].answer.bodyText = Derive.stripQuoted(Derive.htmlToText(html));
        } catch (e) { /* preview text remains the fallback */ }
      }

      var gaps = Derive.findGaps(scanData.topics, scanData.precedents, {});
      var ctx = {
        ownerName: ownerName,
        ownerAddress: scanData.owner,
        generatedOn: new Date().toISOString().slice(0, 10),
        rangeLabel: "last " + scanData.years + " year(s)",
        contentsLabel: "people, open items, topics, attachment index, precedents, exit interview",
        correspondents: scanData.correspondents,
        topics: scanData.topics,
        openWaiting: scanData.open.waiting,
        openOwed: scanData.open.owed,
        attachments: scanData.attachments,
        precedents: scanData.precedents,
        gaps: gaps,
      };
      var plan = Pack.buildPackPlan(ctx);

      var enc = new TextEncoder();
      for (var i = 0; i < plan.length; i++) {
        var f = plan[i];
        setStatus("work", "Writing " + (i + 1) + "/" + plan.length + ": " + f.name);
        var inPrec = f.name.indexOf("precedents/") === 0;
        var parentId = inPrec ? precFolderId : packFolderId;
        var name = inPrec ? f.name.slice("precedents/".length) : f.name;
        var content = f.kind === "docx" ? await docxBuffer(f.doc) : enc.encode(f.text).buffer;
        await GraphData.uploadFile(token, driveId, parentId, name, content);
      }

      // best-effort .agent (SharePoint destinations only — the format needs a site)
      var agentNote = "";
      if (wantAgent && isSharePoint) {
        try {
          setStatus("work", "Writing Ask the Archive.agent…");
          var info = await GraphData.folderInfo(token, driveId, packFolderId);
          var starters = Derive.buildStarters(ownerName, scanData.topics, scanData.correspondents);
          var agentJson = Derive.buildAgentJson({
            ownerName: ownerName, folderUrl: info.webUrl, starters: starters,
          });
          await GraphData.uploadFile(token, driveId, packFolderId,
            "Ask the Archive.agent", enc.encode(JSON.stringify(agentJson, null, 2)).buffer);
          agentNote = " A ready-made agent file was included (best-effort — the manual " +
            "create-an-agent step in README.txt is the supported path).";
        } catch (e) {
          agentNote = " (Agent file skipped: " + ((e && e.message) || e) + ")";
        }
      } else if (wantAgent && !isSharePoint) {
        agentNote = " Agent file skipped — OneDrive destination; move the folder to " +
          "SharePoint and use the README recipe to create the agent.";
      }

      var folder = await GraphData.folderInfo(token, driveId, packFolderId);
      var host = byId("packResult");
      host.hidden = false;
      host.innerHTML = "";
      var p1 = document.createElement("p");
      p1.textContent = "✅ Pack written (" + plan.length + " files)." + agentNote;
      var p2 = document.createElement("p");
      var a = document.createElement("a");
      a.href = folder.webUrl;
      a.target = "_blank";
      a.rel = "noopener";
      a.textContent = "Open the pack folder";
      p2.appendChild(a);
      var p3 = document.createElement("p");
      p3.textContent = "Next: REVIEW every file (delete what shouldn't become institutional " +
        "knowledge), sign manifest.docx, then follow README.txt to create the agent.";
      host.appendChild(p1);
      host.appendChild(p2);
      host.appendChild(p3);
      setStatus("info", "Knowledge pack complete — review before creating any agent.");
    } catch (e) {
      setStatus("error", "Pack build failed: " + ((e && e.message) || e));
    } finally {
      byId("build").disabled = false;
    }
  }
})();
