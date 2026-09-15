# Canvas Bookmarklets

A small set of browser bookmarklets for the [Canvas LMS](https://www.instructure.com/canvas). They bundle into one **Canvas Toolkit** bookmark: click it on any Canvas course page and a menu opens with every tool in it.

**➡️ [Install page](https://jmbenak71.github.io/Canvas_Bookmarklets/)**

## Install

1. Show your bookmarks bar — `Ctrl`+`Shift`+`B` (Windows) or `⌘`+`Shift`+`B` (Mac).
2. Open the [install page](https://jmbenak71.github.io/Canvas_Bookmarklets/) and **drag** the blue *Canvas Toolkit* button onto the bar. Dragging saves it; clicking does not.
3. Open one of your Canvas courses and click the bookmark.

## What's in the toolkit

| Tool | Run it from | What it does |
| --- | --- | --- |
| Syllabus Auto-Builder | A course page | Weekly schedule, assignment list, grading breakdown and unpublished items, built from the course's modules and assignments and downloaded as a Word file. |
| Class Roster Builder | Anywhere (pick a term) | Full student rosters — name, email, section, last access, grade — for all your courses in that term. |
| At-Risk Early Warning Dashboard | Anywhere (pick a term) | Flags inactive, low-grade and missing-work students across a term. |
| Peer Comparison Snapshot | A course page | Anonymized comparison of one student against the rest of the course. |
| Re-Engagement Email Drafter | Anywhere (pick a term) | Drafts check-in emails for at-risk students across a term. |
| Weekly Announcement Auto-Drafter | A course page | Drafts this week's announcement from what is actually due. |
| Discussion to Word Exporter | A discussion page | Exports the thread — original post and all replies — as a .docx file. |

## How it works

Each tool runs in your own browser, signed in as you, and calls the Canvas REST API for the same information Canvas already shows you. There is no server, no API token to paste, and nothing is sent anywhere outside Canvas.

Notes and limits:

- Three tools read the course ID from the page address, so run those from a course page; three ask you to pick a term and work from anywhere; the discussion exporter runs on a discussion.
- Canvas returns at most 100 items per request, so very large courses may need a manual look past that.
- The email and announcement tools write drafts. They do not send or post anything.
- Rosters, grades and engagement data are protected student records. Handle exports accordingly.

## Files

- `index.html` — the install and documentation page (served by GitHub Pages).
- `canvas-toolkit-launcher.html` — the toolkit install page, with the bookmarklet and its full source.
