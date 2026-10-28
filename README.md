# DANSS daily mail-flow report

Every day at 20:00 Baghdad time, GitHub Actions emails tarek.kyyali@friendgroup-iq.com a
report of every message sent and received by diri-alnassir.iq and friendgroup-iq.com
mailboxes in the previous 24 hours (20:00 to 20:00).

- **Email body:** totals, flagged messages, busiest mailboxes.
- **Attachment** `Mail-flow_YYYY-MM-DD.xlsx`: Summary, Mailboxes, Flags, All messages.
- **Flags:** outbound mail to personal webmail, look-alike domains (e.g. `diri-alnassir.com`,
  `gmaij.com`), messages not delivered, and external mail relayed through the tenant.

Only message-trace metadata is used (time, sender, recipient, subject, status, size).
Message bodies and attachments are never read.

## How it signs in

The Entra app **DANSS Daily Mail-Flow Report** (`efcdd68c-7246-4b84-833e-01cf23a456a6`)
trusts this repo's `main` branch through a federated credential, so no password or secret
is stored. It has:

| Permission | Scope |
|---|---|
| Graph `ExchangeMessageTrace.Read.All` | read-only message trace |
| Exchange `Application Mail.Send` | only tarek.kyyali@friendgroup-iq.com (scope "DANSS Report Sender") |

`setup-tenant.ps1` applies these once.

## Settings

Repository variables: `TENANT_ID`, `CLIENT_ID`, `REPORT_SENDER`, `REPORT_TO` (comma-separated
for more recipients), `INTERNAL_DOMAINS`.

Run on demand: **Actions → Daily mail-flow report → Run workflow**. Set `window_end` to
`now` for the last 24 hours up to this minute.

## Turning it off

Disable the workflow in Actions. To revoke access completely, delete the Entra app.
