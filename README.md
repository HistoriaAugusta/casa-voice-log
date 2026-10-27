# CASA Voice Log

A home-screen web app for Cooper Power & Lighting project managers. Record a
voice memo on site; it is posted word for word to the job's Procore daily log
under your own account.

How it works:

1. Sign in once with your Cooper Microsoft account.
2. Type the job number (recent jobs are one tap), record, play back, send.
3. The app mails the memo from your own mailbox to CASA. CASA transcribes it
   and posts it, then emails you a confirmation.
4. Undo in the app (or reply UNDO to the confirmation) removes it.

No signal? The memo is saved on the phone and sends when the signal returns.
Open the app once it is back online if it does not send by itself.

This is a static page. It has no secrets: the Microsoft app and tenant IDs in
`app.js` are public identifiers, and only Cooper accounts can sign in.

Release: bump `VERSION` in `sw.js` with every change so phones pick it up.
