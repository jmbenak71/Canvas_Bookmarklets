# Canvas Bookmarklets

A small set of browser bookmarklets for the [Canvas LMS](https://www.instructure.com/canvas). They bundle into one **Canvas Toolkit** bookmark: click it on any Canvas course page and a menu opens with every tool in it.

**➡️ [Install page](https://jmbenak71.github.io/Canvas_Bookmarklets/)**

## Install

1. Show your bookmarks bar — `Ctrl`+`Shift`+`B` (Windows) or `⌘`+`Shift`+`B` (Mac).
2. Open the [install page](https://jmbenak71.github.io/Canvas_Bookmarklets/) and **drag** the blue *Canvas Toolkit* button onto the bar. Dragging saves it; clicking does not.
3. Open one of your Canvas courses and click the bookmark.

## What's in the toolkit

| Tool | What it does |
| --- | --- |
| Syllabus Auto-Builder | Builds a syllabus summary — weekly schedule, assignment list, grading breakdown, unpublished items — from the course's modules and assignments, and downloads it as a Word file. |
| Class Roster Builder | Builds a printable roster with contact and last-activity information. |
| At-Risk Early Warning Dashboard | Flags students whose grades, submissions or participation suggest they are falling behind. |
| Peer Comparison Snapshot | Compares one student's activity and scores against the rest of the section. |
| Re-Engagement Email Drafter | Drafts check-in messages to students who have gone quiet, for you to review and send. |
| Weekly Announcement Auto-Drafter | Drafts a weekly announcement from what is actually due that week. |
| Discussion to Word Exporter | Exports a discussion thread, posts and replies in order, to a Word document. |
| Engagement Report | Summarizes participation across the course over time. |

## How it works

Each tool runs in your own browser, signed in as you, and calls the Canvas REST API for the same information Canvas already shows you. There is no server, no API token to paste, and nothing is sent anywhere outside Canvas.

Notes and limits:

- Run the toolkit from inside a course — most tools read the course ID from the page address.
- Canvas returns at most 100 items per request, so very large courses may need a manual look past that.
- The email and announcement tools write drafts. They do not send or post anything.
- Rosters, grades and engagement data are protected student records. Handle exports accordingly.

## Files

- `index.html` — the install and documentation page (served by GitHub Pages).
- `canvas-toolkit-launcher.html` — the toolkit install page, with the bookmarklet and its full source.
