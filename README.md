# Rentman LED Equipment Check

This repository contains a standalone GitHub Actions checker that:

1. Reads the watched LED equipment from `data/Export_Equipment_20260715.xlsx`.
2. Looks through upcoming Rentman projects.
3. Checks each project's planned equipment against the spreadsheet.
4. Sends a Slack warning when a matching project is missing the `tec✔️` tag.

The checker uses only Python standard-library modules, so GitHub Actions does not need dependency installation.

## GitHub Setup

Add these repository secrets:

- `RENTMAN_API_TOKEN`: Rentman API token.
- `SLACK_BOT_TOKEN`: Slack bot token with permission to send messages.

Add this repository variable:

- `SLACK_USER_ID`: Slack user ID to message, for example `UFM0TTNHX`.

Alternatively, you can skip `SLACK_BOT_TOKEN` and `SLACK_USER_ID` and add `SLACK_WEBHOOK_URL` as a repository secret for channel webhook delivery instead.

The workflow is in `.github/workflows/rentman-led-check.yml`. It runs daily and can also be started manually from the GitHub Actions tab. Manual runs support a `dry_run` option that prints the Slack message without sending it.

## Matching

By default, the script matches Rentman project equipment exactly against these spreadsheet columns:

- `Code`
- `Name (in database)`
- `General Name`

Useful workflow/script environment variables:

- `MATCH_MODE`: `exact` or `contains`. Use `contains` only if Rentman names do not exactly match the sheet.
- `REQUIRED_TAG`: defaults to `tec✔️`.
- `LOOKAHEAD_DAYS`: leave empty for all future projects, or set a number such as `90`.
- `SLACK_NOTIFY_ON_EMPTY`: set to `true` if you want a Slack confirmation when no projects need warning.
- `RENTMAN_PROJECT_URL_TEMPLATE`: optional Slack link template, for example `https://example.rentmanapp.com/projects/{id}`.

## Local Checks

Preview the parsed equipment targets:

```bash
python scripts/check_rentman_led.py --print-targets
```

Preview the checker without posting to Slack:

```bash
DRY_RUN=true RENTMAN_API_TOKEN=... python scripts/check_rentman_led.py
```
