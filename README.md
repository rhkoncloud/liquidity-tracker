# liquidity-tracker

Monthly recurring-payments liquidity tracker (PWA) synced to my OneDrive workbook.
See SETUP_GUIDE.md to deploy.

## Dashboard tab
Spend-analysis charts (monthly totals vs base, category drill-down, outlier detection,
FY liquidity forward view) built from the "Monthly Spend by Category" and "Liquidity
Calendar" sheets of the same OneDrive workbook, fetched live via Microsoft Graph.
No financial data is stored in this repo or in the service-worker cache.

## App PIN
Optional 4–6 digit PIN lock (Settings ⚙ → Set app PIN). Only a salted SHA-256 hash is
kept in localStorage on the device. Re-locks after 5 minutes in the background.
This deters casual access on an unlocked phone; real protection remains the Microsoft sign-in.
