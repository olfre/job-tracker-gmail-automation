Job tracker: Gmail automation

A Google Apps Script I built for my own job search. It connects Gmail to my Google Sheets application tracker, so replies from companies update the tracker without copy-pasting.

What it does
Checks Gmail once a day (or on demand from a sheet menu) for replies about applications, in English and Polish
Sorts each email into rejection, interview / next step, or confirmation
Rejection matched to exactly one open application: sets the status to Rejected, records the reply date and adds a short note
Interview or next-step email: records the reply date and sets a follow-up date, but leaves the status for me to change
Marks every changed row yellow and logs every email it finds, matched or not, in an "Inbox log" tab with a Gmail link, so nothing changes silently
Labels processed emails so they aren't read twice
Adds a "Job tracker" menu with a quick-add form that warns about repeat applications at the same company
Design choices
Rejection matching is strict on purpose: confirmation emails often contain words like "unfortunately" in their small print, so only the subject and the start of the email count
When several open applications match one company, it picks the role whose words fit the email best; if it can't decide, it asks me to update by hand
Why I built it

Updating the tracker by hand after every email was slow and easy to forget. Now it stays current on its own, and I only check the highlighted rows.

Setup
In your tracker sheet: Extensions > Apps Script, paste Code.gs, save and reload the sheet
Adjust the tab names and column numbers at the top of the script to match your tracker
Run "Job tracker > Turn on daily Gmail check" once and approve the Gmail and Sheets permissions

Built with Google Apps Script, with AI help for parts of the code.
