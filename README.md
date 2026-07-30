# Successor Console

Outlook add-in for the retirement knowledge problem: someone leaves, their
mailbox is handed to a successor, and thirty years of institutional knowledge
is suddenly an unsearchable pile. Successor Console navigates the inherited
mailbox, then distills it into a **knowledge pack** your organization's
Copilot agent can answer from — with citations.

## Two halves

**Navigate (the console).** Scan a mailbox you've been granted access to
(delegated `Mail.Read.Shared` — only works if IT already gave you Full
Access) and browse:

- **People** — who they actually corresponded with, ranked
- **Open items** — asks that never got answered, in both directions: the
  live handoff list
- **Topics** — deterministic topic map of what they worked on
- **Files** — searchable attachment index with links to the carrying emails

**Distill (the knowledge pack).** One click writes a folder of `.docx` files
to OneDrive or SharePoint:

- `manifest.docx` — provenance + a review sign-off line
- `people.docx`, `open-items.docx`, `projects.docx`, `attachments.docx`
- `precedents/<topic>.docx` — **real questions with the person's real
  answers**, verbatim, each linked to its source email
- `Exit interview - <name>.docx` — auto-generated questions for the topics
  the written record is thin on (phone-call tells, thin answers): ask them
  *before* they leave
- `README.txt` — the 30-second create-an-agent recipe + preflights
- `REVIEW-PROMPTS.txt` — optional Copilot prompts that help the human review
- `Ask the Archive.agent` — best-effort ready-made SharePoint agent file
  (community-documented format; the manual recipe is the supported path)

## Design rules (why it's shaped this way)

- **Deterministic extraction, never generation.** Every word in the pack is
  a verbatim copy with a source link. AI enters only at question time, when
  *your* Copilot agent reads the finished pack — so answers cite real
  precedent instead of paraphrasing it. The agent's ground rules (baked into
  the `.agent` file) include: never speak as the person.
- **The pack is a review gate.** A human reads, prunes, and signs before the
  pack becomes agent knowledge — unlike grounding an agent on a raw mailbox.
- **`.docx`, not `.md`** — Markdown isn't usable agent knowledge; each file
  stays small (unlicensed tenants skip SharePoint files over 7 MB).
- **No backend, no AI calls, no data collection.** Delegated Graph only:
  `Mail.Read`, `Mail.Read.Shared`, `Files.ReadWrite.All`.

## Architecture

- `manifest.xml` — add-in-only XML manifest (task pane, Mailbox 1.5)
- `src/derive.js` — pure derivation (correspondents, topics, open items,
  precedent pairing, gap analysis, agent JSON) — Node-tested
- `src/docx.js` — pure minimal OOXML writer (a .docx is a ZIP of XML parts;
  JSZip assembles in the pane) — Node-tested
- `src/pack.js` — pure pack file plan — Node-tested
- `src/graph.js` — Graph data layer, MSAL nested app auth (NAA), no backend
- `src/taskpane/` — the pane UI
- Hosted on GitHub Pages; the manifest points at the hosted files

## Development

```sh
npm test         # offline unit tests (derive + pack + docx)
npm run validate # validate manifest.xml (needs npm i first)
```

Contributions welcome — see [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT © PurposeBuilt Systems
