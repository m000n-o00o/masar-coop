# مسار التعاوني

**Live site:** https://m000n-o00o.github.io/masar-coop/

A guide site for university students going through co-op training: journey map, countdowns, student experiences, courses tracker, CV and LinkedIn guides, and an application tracker.

## Files
- `index.html` — the whole site (single file, no build step).
- `CV_Template.docx` — CV template linked from the CV section. Keep it next to `index.html`.
- `coop-email-tracker.gs` — optional Google Apps Script that logs co-op emails to a Google Sheet. The same code is embedded in the site.

## Notes
- Student data (checklists, countdowns, tracker, courses) is stored in the visitor's browser only (localStorage).
- Student experiences are embedded in `index.html` (`<script id="expData">`). Phone numbers were removed before publishing.
