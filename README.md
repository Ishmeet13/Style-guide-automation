# Assignment 3: Style Guide Automation 
---

## Overview

This is my research submission for the bonus assignment on automating style guide enforcement for financial documents. I've spent the past couple weeks diving deep into different technologies, designing a practical system architecture, and thinking through how this would actually work in production.

The assignment asked for a solution to detect and fix formatting issues in .docx files, so I've put together a comprehensive research document, system design, and some implementation guidance.

---

## Project Structure

```
assignment3/
├── diagrams/
│   ├── Final_Style_guide.png
│   ├── Job Processing Flow (Sequence).png
│   ├── Rule & Document Model (Class Diagram).png
│   ├── processing-workflow-detail.svg
│   └── system-architecture-flowchart.svg
├── docs/
│   ├── SYSTEM_DESIGN.md
│   ├── EXECUTIVE_SUMMARY.md
│   ├── RESEARCH_DOCUMENT.md
│   └── QUICK_START.md
├── examples/
│   └── example_usage.py
├── src/
│   ├── rule_engine.py
│   ├── validator.py
│   ├── corrector.py
│   └── main.py
└── README.md
```

---

## What's Included

I've organized everything into four main documents:

### Documentation

**[RESEARCH_DOCUMENT.md](./docs/RESEARCH_DOCUMENT.md)** (about 30 pages)
This is where I compared different technologies for working with Word documents. I looked at python-docx, Microsoft's Open XML SDK, some commercial options like Aspose, and a bunch of other libraries. For each one, I documented the pros and cons, checked out their performance characteristics, and dug through documentation and Stack Overflow to see what people's real experiences were. I've included links to about 30 different sources - official docs, tutorials, blog posts, and discussions.

**[SYSTEM_DESIGN.md](./docs/SYSTEM_DESIGN.md)** (about 30 pages)
Here's where I mapped out how the system would actually work. I started with the high-level architecture, then broke it down into specific components like the rule engine, validation logic, and correction workflow. I've included detailed diagrams:

- **[System Architecture](./diagrams/Final_Style_guide.png)** - Complete system overview with all components
- **[Job Processing Flow](./diagrams/Job%20Processing%20Flow%20(Sequence).png)** - Sequence diagram showing job lifecycle
- **[Class Diagram](./diagrams/Rule%20%26%20Document%20Model%20(Class%20Diagram).png)** - Data models and relationships
- **[Processing Workflow](./diagrams/processing-workflow-detail.svg)** - Step-by-step processing with timing
- **[Architecture Flowchart](./diagrams/system-architecture-flowchart.svg)** - Detailed data flow

I've also included some pseudocode for the main algorithms and tried to explain the key design decisions - like why I chose a rules-based approach instead of hardcoding everything, and why async processing makes more sense than synchronous.

**[EXECUTIVE_SUMMARY.md](./docs/EXECUTIVE_SUMMARY.md)** (about 20 pages)
This one's more business-focused. I worked out the ROI calculations, compared manual vs automated costs, and put together a realistic implementation timeline. The numbers are pretty compelling - we're looking at about $174K in annual savings with a 2-month break-even period.

**[QUICK_START.md](./docs/QUICK_START.md)** (about 10 pages)
I included this to show what the actual implementation would look like. There's a basic example you could run in 5 minutes, plus a more complete setup guide. I added some troubleshooting tips based on issues I've run into when working with docx files in the past.

### Source Code

The `src/` directory contains working Python code demonstrating the core concepts:

| File | Purpose |
|------|---------|
| `rule_engine.py` | Loads and manages style rules from JSON |
| `validator.py` | Detects formatting violations |
| `corrector.py` | Applies corrections to documents |
| `main.py` | Main orchestrator and CLI interface |

### Examples

The `examples/` directory contains usage demonstrations showing how to use the system for various scenarios.

---

## The Research Process

When I started this, I knew python-docx was a popular choice, but I wanted to make sure I wasn't missing something better. So I spent time looking at alternatives:

I tested out Microsoft's Open XML SDK - it's incredibly powerful and gives you complete control over every aspect of a Word document, but the learning curve is steep and the code gets verbose fast. For this use case, it felt like overkill.

I also looked at Aspose.Words, which is probably the most feature-complete option out there. It's really well-documented and has excellent support, but at $1,199 per developer license, it's hard to justify when open source alternatives can handle what we need.

Apache POI (for Java) and docx4j were interesting if we were committed to a Java stack, but since we're not, python-docx makes more sense. I also checked out PHPWord, LibreOffice's UNO API, and a few conversion tools like Pandoc and Mammoth.js, but those didn't fit the use case as well.

After all that, python-docx came out as the clear winner. It's free, mature, handles everything we need, and has a really clean API. The performance is good enough - we're talking about processing a typical financial report in under 2 minutes, which is totally acceptable.

---

## System Design Approach

The core idea is pretty straightforward: define formatting rules in JSON files, then have the system automatically check documents against those rules and fix any violations.

Here's how it works:

You upload a document through an API endpoint. The system parses the .docx file to extract all the formatting information - fonts, sizes, alignment, table dimensions, everything. Then it loads the relevant rule set and goes through each rule, checking if the document complies. For every violation it finds, it records what's wrong and applies the appropriate fix. At the end, you get back a corrected document and a detailed report showing what was changed.

The rule engine is the interesting part. Instead of hardcoding "company names must be Arial 14pt bold and centered," you define that in a JSON configuration file. This means when the style guide changes (and it always does), someone can update the rules without touching any code. No redeployment, no developer time, just edit the JSON and you're done.

I designed it with a few key components:

The **Rule Engine** loads and parses these JSON rule definitions. It handles versioning so you can track changes over time and roll back if needed.

The **Document Parser** reads the .docx file using python-docx and builds an internal representation of the document structure. This includes all the paragraphs, tables, formatting properties, everything.

The **Validator** takes the parsed document and the rules, then systematically checks each element. It's looking for things like "is this title centered?" or "are the numbers in this column bold?" When it finds a mismatch, it creates a violation record with details about what was expected vs what it found.

The **Corrector** takes those violations and actually fixes them. It modifies the formatting properties while being careful to preserve all the text content. This is important - we're only fixing formatting, not changing what the document says.

The **Reporter** generates a summary of everything that was changed. This gives reviewers confidence that the system did what it was supposed to do.

---

## Why This Design Works

I made a few key decisions that are worth explaining:

**Rules-based vs hardcoded**: I went with rules in JSON files rather than hardcoding the logic. Yes, it's more complex upfront, but think about what happens six months from now when the style guide changes. With hardcoded logic, you'd need a developer to update the code, test it, deploy it. With JSON rules, someone just edits the file. This saves so much time and headache down the road.

**Asynchronous processing**: Documents can take a minute or two to process, especially large ones. If the API call was synchronous, users would be sitting there waiting. Instead, we return immediately with a job ID, process the document in the background, and let them check back or get notified when it's done. This is standard practice for long-running operations and makes for a much better user experience.

**Generating a new document**: I considered modifying the original file in place, but decided against it. By creating a new corrected version, we keep the original intact for audit purposes. This also lets users compare before and after to make sure nothing unexpected happened.

**Using a library instead of automating Office**: We could theoretically spin up actual Word instances and automate them, but that's a nightmare to manage in production. Libraries like python-docx give us everything we need without the overhead and complexity of running GUI applications on a server.

---

## The Business Case

I worked out the numbers and they're pretty compelling. Right now, if someone's spending 30 minutes manually reviewing each document for style guide compliance, that's $25 per document at a typical hourly rate. Multiply that by thousands of documents per year and you're looking at serious money.

With automation, the per-document cost drops to maybe 50 cents (server time, storage, etc.). The time goes from 30 minutes to 2 minutes, because a human just needs to do a quick review of the correction report rather than manually checking everything.

For an organization processing 7,200 documents annually, that's about $180K in current costs vs $6K with automation. The development cost is around $30K for a 6-week project with a small team, so you break even in about 2 months. After that it's pure savings.

Beyond the direct cost savings, there's the quality improvement. Humans get tired and make mistakes - the error rate for manual review is probably in the 5-10% range. An automated system, once properly tested, should be under 1%. That consistency is valuable.

---

## Implementation Considerations

If we were to actually build this, I'd suggest a phased approach:

Start with a basic prototype in the first two weeks - just the core validation and correction logic working with a simple rule set. This proves out the concept and lets us catch any fundamental issues early.

Then spend the next two weeks building out the API, adding async processing with Celery, and setting up the storage infrastructure. This is where we add the production-ready pieces around that core logic.

Week five would be all about testing. Run it against a big batch of real documents, measure the accuracy, find edge cases, fix bugs. This is crucial - you can't skimp on testing when you're automatically modifying people's documents.

Final week is deployment and documentation. Get it running in production, set up monitoring, write the user guides, train the people who'll be using it.

The main risks are around edge cases - there's always some unexpected formatting scenario you didn't anticipate. That's why having good error handling and logging is important. When something fails, you need to know exactly what happened so you can add it to your test suite and fix it.

---

## Technical Details That Matter

A few implementation notes that might not be obvious:

**Performance**: Python-docx is reasonably fast, but if you're processing hundreds of documents in parallel, you'll want to scale horizontally by adding more worker processes. This is straightforward with Celery and Redis.

**Memory management**: Don't try to load massive documents entirely into memory. Python-docx handles this pretty well, but it's something to keep an eye on. For really large files, you might need to process them in chunks.

**Character encoding**: Always fun. .docx files are UTF-8 by default, which is good, but you still need to handle edge cases with special characters properly.

**Table complexity**: Tables are probably the trickiest part. Merged cells, nested tables, varying column widths - there's a lot of complexity there. The test suite needs good coverage of different table scenarios.

**Regex patterns**: For finding things like company names or dates, you'll need some carefully crafted regex patterns. These can get complicated fast, so document them well and test thoroughly.

---

## What I Learned

Going through this research process was interesting. I thought I knew python-docx pretty well, but digging into the alternatives showed me there's a whole ecosystem of tools out there. The Open XML SDK in particular is impressive in its completeness, even if it's overkill for this use case.

I also spent time thinking about the rule engine design. There's a balance between making rules flexible enough to handle everything you might need, and keeping them simple enough that non-developers can work with them. I think JSON hits a good sweet spot there - it's structured enough to be machine-readable but readable enough for humans.

The business case was eye-opening too. I knew automation saves time, but until you actually calculate the numbers, you don't realize just how significant the impact can be. A 93% time reduction is huge.

---

## Resources and References

Throughout this research, I relied on a bunch of different sources. The official python-docx documentation is excellent, and the GitHub issues and Stack Overflow discussions gave me insights into real-world usage patterns and common pitfalls.

I looked at Microsoft's OpenXML documentation for understanding the underlying file format. The Aspose documentation was helpful for seeing what a commercial solution looks like and what features might be worth considering for future enhancements.

Various blog posts and tutorials showed me how people are actually using these tools in production. And academic papers on document automation gave me some theoretical background on the problem space.

All of these sources are linked in the research document, organized by category for easy reference.

---

## Quick Start

To run the basic example:

```bash
# Install dependencies
pip install python-docx

# Run validation
python src/main.py --rules rules.json --document input.docx --output corrected.docx
```

For detailed setup instructions, see [QUICK_START.md](./docs/QUICK_START.md).

---

## Final Thoughts

This was a good assignment to work through. It forced me to think beyond just "here's how you'd code it" and really consider the research process, the design trade-offs, and the business implications.

If I were actually building this, I'd probably start even simpler than what I've outlined - maybe just handle cover page formatting in the first version, prove it works, then gradually add more rules. But having the complete design mapped out gives a clear roadmap for where to go.

The documents I've put together should give you a good sense of how I approach problems - starting with research, considering alternatives, making informed technology choices, and designing solutions that are practical to build and maintain.

Looking forward to discussing this in more detail.

---

*BestCo Technical Work Sample | Assignment 3 | Ishmeet Singh Arora | January 2025*
