/* Offline unit tests for derivation, pack assembly, and the docx writer.
 * Run: npm test */
"use strict";
var D = require("../src/derive.js");
var Docx = require("../src/docx.js");
var Pack = require("../src/pack.js");

var failures = 0;
function check(label, actual, expected) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

/* --------------------------------------------------------------- primitives */

check("subject: stacked prefixes", D.normalizeSubject("RE: Fw: re: Bridge plans"), "bridge plans");
check("subject: whitespace", D.normalizeSubject("  Bridge   plans "), "bridge plans");
check("stripQuoted cuts From:", D.stripQuoted("My answer.\nFrom: Bob\nSent: x\nOld text"), "My answer.");
check("stripQuoted cuts On..wrote:", D.stripQuoted("Yes.\nOn Tue, Jan 2, Ann Lee wrote:\nquoted"), "Yes.");
check("htmlToText basics", D.htmlToText("<p>Hi&nbsp;there</p><div>Second&amp;line</div>"), "Hi there\nSecond&line");
check("asks: question mark", D.asksSomething("x", "Does Friday work?"), true);
check("asks: phrase", D.asksSomething("x", "Please review when you can."), true);
check("asks: statement", D.asksSomething("FYI", "Sharing for the record."), false);

/* ------------------------------------------------------------ correspondents */

function to(addr, name) { return { emailAddress: { address: addr, name: name || addr } }; }
var OWNER = "jane@agency.gov";

var sent = [
  { id: "s1", conversationId: "c1", subject: "Bridge plans", bodyPreview: "Can you review the plans?",
    sentDateTime: "2026-01-10T10:00:00Z", toRecipients: [to("eng@county.gov", "County Engineer")],
    webLink: "https://outlook/s1" },
  { id: "s2", conversationId: "c2", subject: "RE: Permit question", bodyPreview: "The permit number format is X-99. Call me if unclear.",
    sentDateTime: "2026-02-01T10:00:00Z", toRecipients: [to("clerk@city.gov", "City Clerk")],
    webLink: "https://outlook/s2" },
  { id: "s3", conversationId: "c3", subject: "Lunch", bodyPreview: "Tacos?",
    sentDateTime: "2026-02-02T10:00:00Z", toRecipients: [to("pal@agency.gov", "Pal")],
    webLink: "https://outlook/s3" },
];
var inbound = [
  { id: "i1", conversationId: "c2", subject: "Permit question", bodyPreview: "What is the permit number format for driveway permits?",
    receivedDateTime: "2026-01-30T09:00:00Z", from: to("clerk@city.gov", "City Clerk"),
    webLink: "https://outlook/i1" },
  { id: "i2", conversationId: "c9", subject: "Deals!", bodyPreview: "Can you believe these deals? Shop now!",
    receivedDateTime: "2026-01-15T09:00:00Z", from: to("newsletter@shop.com", "Shop"),
    webLink: "https://outlook/i2" },
  { id: "i3", conversationId: "c4", subject: "Culvert sizing", bodyPreview: "How do you size culverts on secondary roads?",
    receivedDateTime: "2026-02-05T09:00:00Z", from: to("eng@county.gov", "County Engineer"),
    webLink: "https://outlook/i3" },
];

var people = D.rankCorrespondents(sent, inbound, OWNER);
check("correspondent count (newsletter excluded)", people.length, 3);
// clerk and engineer tie at 2 emails each; alphabetical tie-break puts clerk first
check("top correspondent (tie-break alphabetical)", people[0].email, "clerk@city.gov");
check("top counts", [people[0].sentTo, people[0].receivedFrom], [1, 1]);
check("name captured", people[0].name, "City Clerk");

/* ----------------------------------------------------------------- topics */

var topics = D.clusterTopics(sent, inbound);
check("RE: groups with original", topics.some(function (t) { return t.key === "permit question" && t.count === 2; }), true);
var permitTopic = topics.filter(function (t) { return t.key === "permit question"; })[0];
check("topic label strips RE:", permitTopic.label, "Permit question");
check("topics sorted by traffic", topics[0].count >= topics[topics.length - 1].count, true);

/* -------------------------------------------------------------- precedents */

var precedents = D.extractPrecedents(sent, inbound, {});
check("one precedent pair (permit Q answered)", precedents.length, 1);
check("precedent question from clerk", precedents[0].question.fromAddress, "clerk@city.gov");
check("precedent answer is s2", precedents[0].answer.id, "s2");
check("tacit flag (“call me” in answer)", precedents[0].tacit, true);
check("precedent keeps links", [precedents[0].question.webLink, precedents[0].answer.webLink],
  ["https://outlook/i1", "https://outlook/s2"]);

// a reply BEFORE the question is not an answer
var precEarly = D.extractPrecedents(
  [{ id: "sx", conversationId: "cz", subject: "RE: q", bodyPreview: "answer",
     sentDateTime: "2026-01-01T00:00:00Z" }],
  [{ id: "ix", conversationId: "cz", subject: "q", bodyPreview: "Can you help?",
     receivedDateTime: "2026-01-02T00:00:00Z", from: to("a@b.gov") }], {});
check("reply before question ignored", precEarly.length, 0);

/* -------------------------------------------------------------- open items */

var open = D.findOpenItems(sent, inbound, {});
// c3 "Tacos?" is a genuine unanswered question — it counts (human reviews the list)
check("waiting: both unanswered asks, oldest first", open.waiting.map(function (w) { return w.conversationId; }), ["c1", "c3"]);
check("waiting who", open.waiting[0].who, "County Engineer");
check("owed: culvert question unanswered (permit was answered)",
  open.owed.map(function (o) { return o.conversationId; }), ["c4"]);
check("owed excludes newsletter", open.owed.some(function (o) { return o.conversationId === "c9"; }), false);

/* -------------------------------------------------------------------- gaps */

var fatTopics = [
  { key: "permit question", label: "Permit question", count: 12, conversations: 9, first: "2024-01-01", last: "2026-02-01" },
  { key: "culvert sizing", label: "Culvert sizing", count: 10, conversations: 8, first: "2024-01-01", last: "2026-02-05" },
  { key: "lunch", label: "Lunch", count: 3, conversations: 3, first: "2026-01-01", last: "2026-02-02" },
];
var gaps = D.findGaps(fatTopics, precedents, { minTraffic: 8 });
check("tacit gap for permit (call-me answer)", gaps.some(function (g) { return g.topic === "Permit question" && g.kind === "tacit"; }), true);
check("thin gap for culverts (no written answer)", gaps.some(function (g) { return g.topic === "Culvert sizing" && g.kind === "thin"; }), true);
check("low-traffic lunch not a gap", gaps.some(function (g) { return g.topic === "Lunch"; }), false);

/* --------------------------------------------------------- agent + starters */

var starters = D.buildStarters("Jane", fatTopics, people);
check("starters count", starters.length, 4);
check("starter cites real topic", starters[0], "How did Jane handle “Permit question”?");

var agent = D.buildAgentJson({ ownerName: "Jane", folderUrl: "https://x.sharepoint.com/sites/s/lib/pack", starters: starters });
check("agent schema version", agent.schemaVersion, "0.2.0");
check("agent name", agent.customCopilotConfig.gptDefinition.name, "Ask Jane's Archive");
check("agent scoped to folder", agent.customCopilotConfig.gptDefinition.capabilities[0].items_by_url[0].url,
  "https://x.sharepoint.com/sites/s/lib/pack");
check("agent anti-persona instruction present",
  agent.customCopilotConfig.gptDefinition.instructions.indexOf("Never speak as Jane") !== -1, true);
check("agent starters carried", agent.customCopilotConfig.conversationStarters.conversationStarterList.length, 4);

/* -------------------------------------------------------------------- pack */

var ctx = {
  ownerName: "Jane Doe", ownerAddress: OWNER, generatedOn: "2026-07-30",
  rangeLabel: "last 3 year(s)", contentsLabel: "everything",
  correspondents: people, topics: topics,
  openWaiting: open.waiting, openOwed: open.owed,
  attachments: [{ name: "plan.pdf", size: 2048, date: "2026-01-10T10:00:00Z", subject: "Bridge plans", from: "County Engineer", webLink: "https://outlook/s1" }],
  precedents: precedents, gaps: gaps,
};
var plan = Pack.buildPackPlan(ctx);
var names = plan.map(function (f) { return f.name; });
check("core files present",
  ["manifest.docx", "people.docx", "open-items.docx", "projects.docx", "attachments.docx", "README.txt", "REVIEW-PROMPTS.txt"]
    .every(function (n) { return names.indexOf(n) !== -1; }),
  true);
check("exit interview named for owner", names.some(function (n) { return n === "Exit interview - Jane Doe.docx"; }), true);
check("one precedent topic file", names.filter(function (n) { return n.indexOf("precedents/") === 0; }).length, 1);
check("precedent file safe-named", names.some(function (n) { return n === "precedents/Permit question.docx"; }), true);
check("README teaches create-an-agent", plan.filter(function (f) { return f.name === "README.txt"; })[0].text.indexOf("Create an agent") !== -1, true);
check("safeName strips illegal chars", Pack._internals.safeName("RE: a/b\\c: d?"), "RE a b c d");

/* -------------------------------------------------------------------- docx */

var built = Docx.buildDocx({
  title: "T & T",
  blocks: [
    { type: "h1", text: "Head <1>" },
    { type: "p", text: "Body & text" },
    { type: "link", text: "Src", url: "https://x.y/z?a=1&b=2" },
    { type: "link", text: "Src2", url: "https://x.y/w" },
    { type: "kv", key: "K", value: "V" },
  ],
});
var docXml = built.parts["word/document.xml"];
var relXml = built.parts["word/_rels/document.xml.rels"];
check("docx has 4 parts", Object.keys(built.parts).length, 4);
check("title escaped", docXml.indexOf("T &amp; T") !== -1, true);
check("heading escaped", docXml.indexOf("Head &lt;1&gt;") !== -1, true);
check("two hyperlink rels", built.relCount, 2);
check("rel ids match doc refs",
  (docXml.match(/r:id="rId100\d"/g) || []).length, 2);
check("rel target escaped", relXml.indexOf("https://x.y/z?a=1&amp;b=2") !== -1, true);
check("external target mode", (relXml.match(/TargetMode="External"/g) || []).length, 2);
check("content types names document part",
  built.parts["[Content_Types].xml"].indexOf("/word/document.xml") !== -1, true);
check("esc strips control chars", Docx._internals.esc("ab"), "ab");

if (failures) {
  console.error("\n" + failures + " test(s) FAILED");
  process.exit(1);
}
console.log("All derive/pack/docx tests passed.");
