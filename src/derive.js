/*
 * Successor Console — pure derivation logic. No Office.js, no Graph, no DOM;
 * unit-testable under Node (see test/derive.test.js).
 *
 * Everything here is DETERMINISTIC EXTRACTION, never generation: the pack's
 * content is verbatim excerpts of what the mailbox owner actually wrote,
 * each carrying a webLink back to the source email. AI (the customer's own
 * Copilot agent) only ever READS the pack at query time — it never writes
 * the record. That division is the product's compliance stance.
 */
(function (root) {
  "use strict";

  var STOPWORDS = ("a an and are as at be by for from has have i in is it of on or our per re " +
    "so that the their this to was we will with you your fw fwd fyi").split(" ");
  var AUTOMATED_RE = /(^|[._-])(no-?reply|donotreply|noreply|notifications?|postmaster|mailer-daemon|newsletter|alerts?)([._-]|@)/i;
  var ASK_PHRASES = [
    "please", "let me know", "can you", "could you", "would you", "will you",
    "do you", "did you", "any update", "thoughts on", "what is", "when is",
    "how do", "need your", "waiting on", "advise", "confirm",
  ];
  // phrases that suggest the real answer happened OFF the record — the
  // gap-detection tell that feeds the exit interview
  var TACIT_RE = /\b(as (we )?discussed|per our (call|conversation|phone)|call me|gave you a call|talked (about|through) (this|it)|on the phone|we spoke)\b/i;

  function norm(addr) { return String(addr || "").trim().toLowerCase(); }

  function normalizeSubject(subject) {
    var s = String(subject || "").trim();
    var prev = null;
    while (prev !== s) {
      prev = s;
      s = s.replace(/^\s*(re|fw|fwd|aw|sv)\s*:\s*/i, "");
    }
    return s.replace(/\s+/g, " ").trim().toLowerCase();
  }

  /** Owner's own words: cut the quoted thread off a body/preview. */
  function stripQuoted(text) {
    var t = String(text || "");
    var cutters = [/\r?\nFrom:\s/i, /\r?\n-+\s*Original Message\s*-+/i, /\r?\nOn .{5,80} wrote:/i, /\r?\n>+\s/];
    var cut = t.length;
    cutters.forEach(function (re) {
      var m = t.search(re);
      if (m !== -1 && m < cut) { cut = m; }
    });
    return t.slice(0, cut).trim();
  }

  function htmlToText(html) {
    return String(html || "")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div|li|tr|h[1-6])>/gi, "\n")
      .replace(/<[^>]+>/g, " ")
      .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&")
      .replace(/&#39;/g, "'").replace(/&quot;/g, '"').replace(/&nbsp;/g, " ")
      .replace(/[ \t]+/g, " ").replace(/[ \t]*\n[ \t]*/g, "\n")
      .replace(/\n{3,}/g, "\n\n").trim();
  }

  function asksSomething(subject, ownText) {
    var t = (String(subject || "") + " " + String(ownText || "")).toLowerCase();
    if (ownText && ownText.indexOf("?") !== -1) { return true; }
    return ASK_PHRASES.some(function (p) {
      return new RegExp("\\b" + p.replace(/ /g, "\\s+") + "\\b", "i").test(t);
    });
  }

  /* ----------------------------------------------------------- correspondents */

  /**
   * Who did the owner actually deal with? Counts exchanges per counterpart:
   * recipients of the owner's sent mail + senders of inbound. Automated
   * senders excluded. Sorted most-dealt-with first.
   */
  function rankCorrespondents(sent, inbound, ownerAddress) {
    var me = norm(ownerAddress);
    var people = {}; // email -> {name, sentTo, receivedFrom, first, last}
    function touch(email, name, dateIso, field) {
      var e = norm(email);
      if (!e || e === me || AUTOMATED_RE.test(e)) { return; }
      var p = people[e] || (people[e] = { email: e, name: "", sentTo: 0, receivedFrom: 0, first: dateIso, last: dateIso });
      p[field]++;
      if (name && name !== e && name.length > p.name.length) { p.name = name; }
      if (dateIso < p.first) { p.first = dateIso; }
      if (dateIso > p.last) { p.last = dateIso; }
    }
    (sent || []).forEach(function (m) {
      (m.toRecipients || []).concat(m.ccRecipients || []).forEach(function (r) {
        var ea = r.emailAddress || {};
        touch(ea.address, ea.name, m.sentDateTime, "sentTo");
      });
    });
    (inbound || []).forEach(function (m) {
      var ea = (m.from && m.from.emailAddress) || {};
      touch(ea.address, ea.name, m.receivedDateTime, "receivedFrom");
    });
    return Object.keys(people).map(function (e) {
      var p = people[e];
      p.total = p.sentTo + p.receivedFrom;
      p.name = p.name || e;
      return p;
    }).sort(function (a, b) { return b.total - a.total || a.email.localeCompare(b.email); });
  }

  /* ----------------------------------------------------------------- topics */

  /**
   * Deterministic topic clustering: conversations sharing a normalized
   * subject are one topic. Label = shortest original subject; keywords =
   * most frequent non-stopword tokens (used for conversation starters).
   */
  function clusterTopics(sent, inbound) {
    var topics = {}; // key -> {label, count, convIds:{}, keywords:{}, first, last}
    function add(msg, dateIso) {
      var key = normalizeSubject(msg.subject);
      if (!key) { key = "(no subject)"; }
      var t = topics[key] || (topics[key] = {
        key: key, label: String(msg.subject || "(no subject)").replace(/^\s*((re|fw|fwd)\s*:\s*)+/i, "").trim() || "(no subject)",
        count: 0, convIds: {}, first: dateIso, last: dateIso,
      });
      t.count++;
      if (msg.conversationId) { t.convIds[msg.conversationId] = true; }
      var label = String(msg.subject || "").replace(/^\s*((re|fw|fwd)\s*:\s*)+/i, "").trim();
      if (label && label.length < t.label.length) { t.label = label; }
      if (dateIso < t.first) { t.first = dateIso; }
      if (dateIso > t.last) { t.last = dateIso; }
    }
    (sent || []).forEach(function (m) { add(m, m.sentDateTime || ""); });
    (inbound || []).forEach(function (m) { add(m, m.receivedDateTime || ""); });
    var list = Object.keys(topics).map(function (k) {
      var t = topics[k];
      t.conversations = Object.keys(t.convIds).length;
      delete t.convIds;
      return t;
    }).sort(function (a, b) { return b.count - a.count || a.key.localeCompare(b.key); });
    return list;
  }

  function keywords(text, max) {
    var counts = {};
    String(text || "").toLowerCase().replace(/[^a-z0-9\s'-]/g, " ").split(/\s+/).forEach(function (w) {
      if (w.length < 4 || STOPWORDS.indexOf(w) !== -1) { return; }
      counts[w] = (counts[w] || 0) + 1;
    });
    return Object.keys(counts)
      .sort(function (a, b) { return counts[b] - counts[a] || a.localeCompare(b); })
      .slice(0, max || 5);
  }

  /* -------------------------------------------------------------- precedents */

  /**
   * The pack's crown jewels: real question -> the owner's real answer.
   * An inbound message that asks something, followed by the owner's reply on
   * the same conversation, becomes a precedent pair. Verbatim, with links.
   *
   * Bodies are optional: pairs are built from previews, then the caller can
   * hydrate `answer.bodyText` for the selected pairs (Graph fetch) — the
   * docx renderer prefers bodyText when present.
   */
  function extractPrecedents(sent, inbound, opts) {
    opts = opts || {};
    var maxPairs = opts.maxPairs || 200;
    var sentByConv = {};
    (sent || []).forEach(function (m) {
      (sentByConv[m.conversationId] = sentByConv[m.conversationId] || []).push(m);
    });
    Object.keys(sentByConv).forEach(function (k) {
      sentByConv[k].sort(function (a, b) { return a.sentDateTime < b.sentDateTime ? -1 : 1; });
    });
    var pairs = [];
    (inbound || []).forEach(function (q) {
      var ea = (q.from && q.from.emailAddress) || {};
      if (AUTOMATED_RE.test(norm(ea.address))) { return; }
      var own = stripQuoted(q.bodyPreview || "");
      if (!asksSomething(q.subject, own)) { return; }
      var replies = sentByConv[q.conversationId] || [];
      var answer = null;
      for (var i = 0; i < replies.length; i++) {
        if (replies[i].sentDateTime > q.receivedDateTime) { answer = replies[i]; break; }
      }
      if (!answer) { return; }
      pairs.push({
        topicKey: normalizeSubject(q.subject),
        subject: String(q.subject || "").replace(/^\s*((re|fw|fwd)\s*:\s*)+/i, "").trim() || "(no subject)",
        question: {
          id: q.id, from: ea.name || ea.address || "", fromAddress: norm(ea.address),
          date: q.receivedDateTime, text: own, webLink: q.webLink || "",
        },
        answer: {
          id: answer.id, date: answer.sentDateTime,
          text: stripQuoted(answer.bodyPreview || ""), webLink: answer.webLink || "",
        },
        tacit: TACIT_RE.test(own) || TACIT_RE.test(stripQuoted(answer.bodyPreview || "")),
      });
    });
    // newest first within a topic, topics with more pairs first
    var byTopic = {};
    pairs.forEach(function (p) { (byTopic[p.topicKey] = byTopic[p.topicKey] || []).push(p); });
    var ordered = [];
    Object.keys(byTopic)
      .sort(function (a, b) { return byTopic[b].length - byTopic[a].length || a.localeCompare(b); })
      .forEach(function (k) {
        byTopic[k].sort(function (a, b) { return a.question.date < b.question.date ? 1 : -1; });
        ordered = ordered.concat(byTopic[k]);
      });
    return ordered.slice(0, maxPairs);
  }

  /* -------------------------------------------------------------- open items */

  /**
   * The live handoff list: asks the owner sent that never got a reply, and
   * inbound asks the mailbox never answered. Departure context — no minimum
   * age, everything still open at scan time counts. Newest ask per
   * conversation wins; automated counterparts excluded.
   */
  function findOpenItems(sent, inbound, opts) {
    opts = opts || {};
    var cap = opts.cap || 100;
    var inByConv = {}, sentByConv = {};
    (inbound || []).forEach(function (m) {
      (inByConv[m.conversationId] = inByConv[m.conversationId] || []).push(m.receivedDateTime);
    });
    (sent || []).forEach(function (m) {
      (sentByConv[m.conversationId] = sentByConv[m.conversationId] || []).push(m.sentDateTime);
    });

    var waiting = {};
    (sent || []).forEach(function (m) {
      var rcpts = (m.toRecipients || []);
      if (!rcpts.length) { return; }
      var first = rcpts[0].emailAddress || {};
      if (AUTOMATED_RE.test(norm(first.address))) { return; }
      var own = stripQuoted(m.bodyPreview || "");
      if (!asksSomething(m.subject, own)) { return; }
      var answered = (inByConv[m.conversationId] || []).some(function (d) { return d > m.sentDateTime; });
      if (answered) { return; }
      var prev = waiting[m.conversationId];
      if (!prev || m.sentDateTime > prev.date) {
        waiting[m.conversationId] = {
          conversationId: m.conversationId, subject: m.subject || "(no subject)",
          date: m.sentDateTime, who: first.name || first.address || "",
          text: own, webLink: m.webLink || "",
        };
      }
    });

    var owed = {};
    (inbound || []).forEach(function (m) {
      var ea = (m.from && m.from.emailAddress) || {};
      if (AUTOMATED_RE.test(norm(ea.address))) { return; }
      var own = stripQuoted(m.bodyPreview || "");
      if (!asksSomething(m.subject, own)) { return; }
      var answered = (sentByConv[m.conversationId] || []).some(function (d) { return d > m.receivedDateTime; });
      if (answered) { return; }
      var prev = owed[m.conversationId];
      if (!prev || m.receivedDateTime > prev.date) {
        owed[m.conversationId] = {
          conversationId: m.conversationId, subject: m.subject || "(no subject)",
          date: m.receivedDateTime, who: ea.name || ea.address || "",
          text: own, webLink: m.webLink || "",
        };
      }
    });

    function toList(map) {
      return Object.keys(map).map(function (k) { return map[k]; })
        .sort(function (a, b) { return a.date < b.date ? -1 : 1; })
        .slice(0, cap);
    }
    return { waiting: toList(waiting), owed: toList(owed) };
  }

  /* ------------------------------------------------------------------- gaps */

  /**
   * Thin spots feed the exit interview: topics with real traffic but few
   * written answers, and precedents whose text points at phone calls.
   * Returns question prompts a human asks the retiree — pure gap analysis,
   * nothing generated.
   */
  function findGaps(topics, precedents, opts) {
    opts = opts || {};
    var minTraffic = opts.minTraffic || 8;
    var answered = {};
    var tacitTopics = {};
    (precedents || []).forEach(function (p) {
      answered[p.topicKey] = (answered[p.topicKey] || 0) + 1;
      if (p.tacit) { tacitTopics[p.topicKey] = p; }
    });
    var gaps = [];
    (topics || []).forEach(function (t) {
      if (t.count < minTraffic) { return; }
      var n = answered[t.key] || 0;
      if (tacitTopics[t.key]) {
        gaps.push({
          topic: t.label, kind: "tacit",
          prompt: "Emails about “" + t.label + "” refer to phone calls or in-person discussions. " +
            "What was decided in those conversations, and what should your successor know that was never written down?",
        });
      } else if (n < 2) {
        gaps.push({
          topic: t.label, kind: "thin",
          prompt: "“" + t.label + "” shows " + t.count + " emails but few written answers from you. " +
            "Describe the process, the key contacts, and anything a successor would not find in the file.",
        });
      }
    });
    return gaps;
  }

  /* ----------------------------------------------------- agent + starters */

  /** Conversation starters derived from what the archive actually contains. */
  function buildStarters(ownerName, topics, correspondents) {
    var starters = [];
    (topics || []).slice(0, 2).forEach(function (t) {
      starters.push("How did " + ownerName + " handle “" + t.label + "”?");
    });
    if (correspondents && correspondents[0]) {
      starters.push("Who did " + ownerName + " work with most, and about what?");
    }
    starters.push("What was still open when " + ownerName + " left?");
    return starters.slice(0, 4);
  }

  /**
   * Best-effort SharePoint `.agent` file (community-documented schema 0.2.0 —
   * not officially specified; the manual create-an-agent recipe in README.txt
   * is the supported path). Instructions bake in the anti-persona guardrail.
   */
  function buildAgentJson(opts) {
    return {
      customCopilotConfig: {
        conversationStarters: {
          conversationStarterList: (opts.starters || []).map(function (s) { return { text: s }; }),
          welcomeMessage: {
            text: "This agent answers from " + opts.ownerName + "'s reviewed knowledge pack. " +
              "Every answer cites its source email.",
          },
        },
        gptDefinition: {
          name: "Ask " + opts.ownerName + "'s Archive",
          description: "Answers questions from the reviewed knowledge pack derived from " +
            opts.ownerName + "'s mailbox. Citations link to the source emails.",
          instructions:
            "Answer ONLY from the knowledge files in this folder. Always cite the source " +
            "email link shown next to the excerpt you used. If the answer is not in the " +
            "archive, say plainly that it is not in the archive. Never speak as " +
            opts.ownerName + " or in the first person; describe what the archive shows. " +
            "These files are verbatim excerpts of real correspondence — do not extrapolate " +
            "beyond them.",
          capabilities: [{
            name: "OneDriveAndSharePoint",
            items_by_sharepoint_ids: [],
            items_by_url: opts.folderUrl ? [{ url: opts.folderUrl }] : [],
          }],
        },
      },
      schemaVersion: "0.2.0",
    };
  }

  root.Derive = {
    normalizeSubject: normalizeSubject,
    stripQuoted: stripQuoted,
    htmlToText: htmlToText,
    asksSomething: asksSomething,
    rankCorrespondents: rankCorrespondents,
    clusterTopics: clusterTopics,
    keywords: keywords,
    extractPrecedents: extractPrecedents,
    findOpenItems: findOpenItems,
    findGaps: findGaps,
    buildStarters: buildStarters,
    buildAgentJson: buildAgentJson,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = root.Derive; }
})(typeof self !== "undefined" ? self : this);
