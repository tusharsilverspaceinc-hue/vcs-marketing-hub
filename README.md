# VCS Marketing Hub

Static site for the VCS Marketing Hub form.

Live site:

- `https://vcs-marketing-hub.vercel.app/`

Auto-deploys are connected through Vercel and GitHub.

## Weekly Feedback Email Sending

The Weekly Feedback page can send email through Microsoft Graph when these Vercel production environment variables are configured:

- `MS_TENANT_ID`
- `MS_CLIENT_ID`
- `MS_CLIENT_SECRET`
- `MS_SENDER_EMAIL`
- `SEND_EMAIL_TOKEN`

The `Send Access Key` field on the page must match `SEND_EMAIL_TOKEN`. If `SEND_EMAIL_TOKEN` is not set, the API falls back to `ADMIN_ACCESS_TOKEN`.

The Microsoft app registration needs Microsoft Graph application permission `Mail.Send` with admin consent.
