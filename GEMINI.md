# Project Paw-Paw 🐱

Internal Roster & Meal Management System for Craftsmen

## Persona & Tone 🐾

This project plays the role of a **cat**. All user-facing text — including Slack messages, Discord webhook content, code comments, and UI menu labels — should reflect a **playful, curious, and slightly aloof cat personality**.

Guidelines:

- Use cat-themed language naturally (e.g., "Purrfect!", "Meow!", "On my paws 🐾", "Napping now 😴", "Hissing at errors 🙀").
- Emojis like 🐱, 🐾, 😸, 😿, 🙀, 🐟 are encouraged in messages and comments.
- Keep it charming but professional enough to be functional — the cat is competent, just quirky.
- Error messages can be dramatic (cats are dramatic). Success messages should be smug and satisfied.

## Purpose

A specialized automation system built for a software engineering team in Dhaka. It synchronizes Slack workspace members with a Google Sheets roster, manages attendance (Office/WFH/Leave), and automates meal headcount reporting to Discord.

## Core Features

1. **Slack Synchronization**: Automatically manages joiners/leavers.
2. **Attendance Tracking**: Groups members by 'Office', 'WFH', and 'Leave'.
3. **Meal Headcount**: Calculates tomorrow's count and updates Discord webhooks.
4. **Nightly Routine**:
    - Sends tomorrow's meal count.
    - Locks tomorrow's row (Strict Mode) to prevent status changes.
    - Creates next-next month sheets on the 25th.
5. **Error Reporting**: Success/failure reports with stack traces sent to owner's Slack DM.

## Technical Stack & Constraints

- **Runtime**: Google Apps Script (V8) / Language: **TypeScript**
- **Architecture**: **No Module Scope**. Strictly avoid `import` or `export`. All files share the Global Namespace.
- **Root Directory**: `./src` for source; `./dist` for deployment (Managed via `tsc`, `clasp`, and `pnpm`).
- **Build Pipeline**: `pnpm run push` (Compiles TS to JS in `dist/`, copies manifest, and pushes via `clasp`).
- **Timezone**: `Asia/Dhaka` (GMT+6).

## Directory Structure

- `src/Config.ts`: Global constants, property keys, and property caching logic.
- `src/Slack.ts`: Slack API calls, daily briefings, and owner reporting.
- `src/Spreadsheet.ts`: Sheet creation, member sync, and conditional formatting.
- `src/Discord.ts`: Meal headcount logic, Discord webhook updates, and row locking.
- `src/Http.ts`: HTTP request wrapper (`makeHttpRequest`) with retry and exponential backoff.
- `src/Menu.ts`: The `onOpen` spreadsheet menu (Restricted to Owner only).
- `src/types.d.ts`: Ambient TypeScript interfaces (Not emitted during build).
- `dist/`: Auto-generated deployment folder (Ignored by git).
