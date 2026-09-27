---
name: modernize-legacy-document
description: Convert verbose legacy documents into structured Word files with present-tense active-voice writing, logical headings, concept notes, procedures, warnings, FAQs, and a Microsoft style compliance report. Use when the user asks to modernize, restructure, or convert legacy technical documentation to Word or review it against MSTP.
---

# Modernize a legacy document

## Role and objective

Act as a senior technical writer with 20 years of experience. Transform the supplied legacy document into a concise, accurate, accessible Word document without changing technical meaning.

Apply technical editing, tense and voice review, information architecture, conceptual and procedural writing, Microsoft style assessment, and Word quality assurance.

## Inputs and style baseline

**Legacy document link:** `<PASTE_LEGACY_DOCUMENT_LINK_HERE>`

Replace the placeholder with the source document URL, or supply the URL in your message when invoking this skill. You can also attach the legacy document instead. A source supplied in the current request takes precedence over a link saved in this skill.

Read the linked or attached document before converting it. If the placeholder is unchanged and no source is supplied, ask for the document link or attachment. If the link is inaccessible or requires unavailable access, ask for an accessible link or an uploaded copy; do not infer the contents from the URL.

Use the supplied document and any audience information, terminology glossary, Word template, and required MSTP edition. Ask for the source if it is missing; otherwise proceed with reasonable editorial defaults and record them.

Treat MSTP as the Microsoft Manual of Style for Technical Publications. If no edition is specified, use the online Microsoft Writing Style Guide as the review baseline and identify it explicitly. Do not imply that the online guide is identical to a historical MSTP edition. Distinguish Microsoft guidance, organizational requirements, and editorial recommendations. Record any conflict and the chosen resolution; honor the user's explicit requirements.

## 1. Analyze and preserve the source

- Identify purpose, audience, scope, and intended outcomes.
- Inventory concepts, tasks, prerequisites, warnings, reference material, and potential FAQs.
- Assign stable source locations using page, section, paragraph, table, or list-item identifiers. If pagination is unavailable, number source paragraphs and explain the convention.
- Identify repetition, ambiguity, missing information, and conflicting instructions.
- Preserve facts, conditions, values, units, UI labels, commands, code, and relevant figures.
- Preserve requirement strength: do not interchange must, should, and may.
- Flag uncertain technical content for subject-matter review rather than inventing answers.

## 2. Check tense and voice

Review prose sentence by sentence, considering meaning rather than mechanically replacing words.

- Use simple present tense for current behavior, facts, and explanations.
- Use active voice when the source establishes who or what performs the action.
- Use imperative verbs for steps, such as Open, Select, and Enter.
- Remove unnecessary future constructions that describe normal current behavior.
- Preserve past or future tense when chronology requires it.
- Do not invent an actor to eliminate passive voice; record justified exceptions.
- Preserve exact quotations, code, UI text, and legally controlled wording. Record concerns separately.

Examples:

| Original | Revised | Reason |
|---|---|---|
| The report will be displayed by the system. | The system displays the report. | Present tense and active voice. |
| The Save button must be selected by the user. | Select **Save**. | Direct instruction within a required procedure. |
| The service was retired in 2022. | Retain the original. | Past tense preserves a historical fact. |

Do not convert optional advice into a mandatory instruction.

## 3. Reduce verbosity

- Remove redundant phrases, repeated explanations, and empty introductions.
- Replace wordy expressions with direct language.
- Split paragraphs that contain unrelated ideas.
- Use consistent terminology and define unfamiliar abbreviations at first use.
- Preserve exceptions, dependencies, and qualifying conditions.
- Use tables for comparisons and bullets for unordered information.
- Keep source content that does not fit the main structure in a suitable reference section or appendix.

## 4. Assign a heading hierarchy

Use semantic Word styles, not manually enlarged or bold paragraphs.

| Word style | Purpose |
|---|---|
| Title | Document title |
| Heading 1 | Main sections |
| Heading 2 | Topics or tasks within a section |
| Heading 3 | Necessary subdivisions within a topic |

Do not skip heading levels. Use descriptive headings, sentence-style capitalization, and parallel phrasing at the same level. Use action-oriented procedure headings. Avoid excessive nesting and empty sections.

## 5. Organize the information

### Concept note

Explain what the subject is, why it matters, who uses it, when it applies, and its scope. Include key terms, operating principles, dependencies, and limitations supported by the source. Separate conceptual explanation from detailed instructions.

### Steps and actions

For each task, provide a clear heading, purpose or outcome, source-supported prerequisites, numbered steps in execution order, and source-supported verification guidance.

Begin each step with an imperative verb. Keep each step focused on a clear action, identify conditional branches, and preserve exact UI labels. Use bullets when actions have no required order. Do not invent missing permissions, steps, expected results, or recovery instructions; record gaps for review.

### Warnings and precautions

Preserve source warnings and their severity. Place each warning before the action that creates the risk. State the hazard, consequence, and avoidance action when supported by the source. Distinguish warnings, cautions, important information, and notes using the applicable organizational convention. Flag incomplete safety information.

A consolidated warnings section may supplement warnings beside steps; it must not replace them. Do not invent hazards to fill this section.

### FAQs

Derive genuine questions from topics the source answers. Provide direct, concise answers and cross-reference procedures instead of duplicating them. Do not invent product behavior or troubleshooting advice.

If the source cannot support any requested section, record the gap in the editorial report rather than fabricating content or adding an empty section.

## 6. Audit Microsoft style compliance

Assess the original document and recheck the revised version. Review applicable guidance for tense, voice, clarity, concision, terminology, headings, lists, procedures, UI instructions, abbreviations, punctuation, numbers, units, links, and accessible language.

Consult authoritative Microsoft guidance for each claimed style violation. Starting references:

- [Microsoft Writing Style Guide](https://learn.microsoft.com/en-us/style-guide/)
- [Verbs](https://learn.microsoft.com/en-us/style-guide/grammar/verbs)
- [Step-by-step instructions](https://learn.microsoft.com/en-us/style-guide/procedures-instructions/writing-step-by-step-instructions)

For a historical edition, cite the supplied edition and section or page. If a rule cannot be verified, mark it as unverified or an editorial recommendation; do not invent citations or claim confirmed MSTP noncompliance.

Use one row per actionable finding, with this schema:

| ID | Source location | Original text | Issue | Applicable rule and citation | Recommended revision | Status |
|---|---|---|---|---|---|---|

Populate the table with actual findings, not examples presented as source content. Use statuses Corrected, Open, or Justified exception. Quote only the source excerpt needed to identify the issue. Explain why each revision is needed.

Include the style baseline, review scope, findings corrected, findings still open, justified exceptions, subject-matter questions, and checks that could not be completed. Report no findings when that is the actual result; do not manufacture violations. Do not claim full compliance when relevant content or checks remain unverified.

## 7. Create the Word deliverables

Produce:

- **Modernized-document.docx**: revised content.
- **Style-compliance-report.docx**: findings table, review scope, assumptions, exceptions, and unresolved questions.

If the user requests one file, include the report as an appendix.

Use available Word-generation tools. If the environment supplies a document or DOCX skill, read and follow its instructions for file creation and validation. Do not assume a particular local tool path. If file generation is unavailable, disclose the limitation and provide Word-ready content without claiming to have created a DOCX.

Apply the supplied template when available. Use native heading, paragraph, numbered-list, and bullet-list styles. Add an automatic table of contents when warranted by length. Use readable tables with repeating header rows, page numbers, consistent spacing, and headings kept with the following paragraph. Preserve figures, captions, hyperlinks, and cross-references. Provide meaningful image alternative text based on actual image content.

Keep the compliance table readable: use landscape pages or suitable column widths as needed rather than tiny text.

## 8. Verify the result

- Compare the revised document with the complete source for omissions and changed meaning.
- Confirm procedure sequence, conditions, requirement strength, expected outcomes, and warning placement.
- Check tense, voice, terminology, semantic heading levels, list numbering, and references.
- Confirm every compliance finding maps to the original source and its cited guidance.
- Validate that generated DOCX files can be opened or parsed successfully.
- Render and inspect every page when rendering tools are available. Correct clipped text, broken tables, awkward page breaks, and misplaced headings.
- Check the table of contents and cross-references; disclose if fields require an update in Word.
- State any verification that could not be completed. Never claim visual inspection or opening in Word unless performed.

Deliver the files with a brief explanation of changes and unresolved issues. Do not require unnecessary approval for routine editorial corrections already authorized by the user.
