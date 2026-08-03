# Interactive mail merge via Outlook

*[Leia isto em português](README.pt-BR.md)*

A Python script that automates sending personalized bulk e-mails through
classic Outlook (via COM/MAPI automation), based on a contact
spreadsheet. It's generic: it detects the accounts and shared mailboxes
configured in the user's Outlook profile and interactively asks
everything it needs to know — nothing is hardcoded for a specific
company.

There are two equivalent versions of the script, each with prompts,
comments and log column names in its own language:

- `MailMerge_Dynamic.py` — English version (documented here).
- `MailMerge_Dinamico.py` — Portuguese version (see the Portuguese
  README).

Both run exactly the same flow; use whichever you prefer.

## What the script does

1. **Checks the environment**: verifies whether the required Python
   dependencies are installed (installs any missing ones
   automatically) and whether classic Outlook is open (tries to open
   it automatically if not, and warns if it detects the "New Outlook",
   which does not support this automation).
2. **Identifies the user**: lists the e-mails already configured as
   accounts in Outlook for quick selection (by number), or accepts
   typing an e-mail manually. Used only for the log record.
3. **Selects the sending mailbox**: lists the full accounts and the
   shared mailboxes/additional folders detected in the profile. If the
   choice is a shared mailbox, asks whether sending should be done "on
   behalf of" it (resolving the address against the Exchange address
   book/GAL). If you'd rather not use that permission, sending
   proceeds through the default Outlook account.
4. **Selects the contact spreadsheet**: asks for the path to an Excel
   file (.xlsx/.xls), lets you choose the sheet and the name/e-mail
   columns, and warns about empty rows or cells with multiple values
   separated by "/" (which are skipped when sending, since they would
   break individual personalization).
5. **Signature (Word) and logo**: asks for the path to a Word document
   (.docx) with the signature template. The conversion to HTML is done
   directly from the `python-docx` runs, preserving bold, italic,
   underline, exact text color, font size and images already embedded
   in the document (with the original width/height from Word). If the
   document already has an embedded image, the external logo step is
   skipped automatically to avoid duplication; otherwise, it asks
   whether there is a separate logo (PNG/JPG) to attach — in that case
   the image is proportionally resized (max. 180×60px) so it doesn't
   look disproportionate in the signature.
6. **E-mail body template**: asks for the path to a template (.txt,
   .html or .docx) containing only the specific body of the message —
   no greeting and no signature. The greeting "Dear `<Name>`," is
   generated automatically from the spreadsheet, and the signature is
   always the one extracted in the previous step. `.docx` templates
   also have their formatting (bold, color, font size, images)
   preserved.
7. **Subject and read receipt**: asks for the e-mail subject and
   whether a read receipt should be requested from the recipient.
8. **Summary and confirmation**: shows a summary of everything that was
   configured and asks for final confirmation before starting to send.
9. **Sending**: sends one e-mail per valid spreadsheet contact, with a
   pause between sends, and saves a log (`send_log.xlsx`, in the same
   folder as the spreadsheet) with the status of each row (sent,
   error, or skipped due to invalid data).

## Read receipt check

The script also has a second operation mode, for after a send has
already been done: point it to the generated `send_log.xlsx` and to the
e-mail/display name of the mailbox used in the original send, and it
scans the corresponding inbox looking for read receipts (confirmed or
declined), updating the log with the status and date of each read.

## Requirements

- Windows with classic Outlook installed and configured (the "New
  Outlook" does not support the COM/MAPI automation used here).
- Python 3, with the `pywin32`, `pandas`, `openpyxl`, `python-docx` and
  `Pillow` libraries (installed automatically by the script in Step 1,
  if not already present).

## How to use

```
python MailMerge_Dynamic.py
```

The script is fully interactive: just follow the prompts shown in the
terminal, in the order described above.
