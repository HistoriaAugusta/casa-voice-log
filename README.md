# CASA Voice Log

A home-screen web app for Cooper Power & Lighting project managers. Record a
voice memo on site; it is posted word for word to the job's Procore daily log
under your own account.

How it works:

1. Sign in with your Cooper Microsoft account. The sign-in lasts about a day;
   after that the app takes you to Microsoft by itself.
2. Pick the job from the list, or type the number. Record, play back, then
   send, re-record or delete.
3. The app mails the memo from your own mailbox to CASA. CASA transcribes it
   and posts it, then emails you a confirmation.
4. Undo in the app (or reply UNDO to the confirmation) removes it.

No signal? The memo is saved on the phone and sends when the signal returns,
while the app is open. It is deleted from the phone once it is sent.

The job list comes from CASA through a hidden folder in your own mailbox, so
the app never needs to reach CASA directly.

This is a static page. It has no secrets: the Microsoft app and tenant IDs in
`app.js` are public identifiers, and only Cooper accounts can sign in.

Release: bump `VERSION` in `sw.js` and `APP_VERSION` in `app.js` together with
every change. The version shows at the bottom of the page.
