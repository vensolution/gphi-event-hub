# Event Hub (GitHub Pages)

One static site: HR Forms, Inventory, Assets/Links and Event Games (Spin the Wheel, Game Tracker, Game Admin). Sign-in and saved data use your existing Firebase project, so nothing changes in your data.

## Put it on GitHub Pages
1. Create a GitHub repository and upload everything in this folder (keep the `bybet` folder and the `.nojekyll` file).
2. Repository **Settings > Pages**: Source = **Deploy from a branch**, Branch = **main**, Folder = **/ (root)**. Save.
3. Your address will be `https://<your-username>.github.io/<repo-name>/`.
4. In Firebase: **Authentication > Settings > Authorized domains > Add domain**, and add `<your-username>.github.io`.
5. Firestore rules: if not published yet, paste `firestore.rules` in Firebase and Publish. They are what protect the data.
6. Check `firebase-config.js` matches your Firebase project.

## Event Games
- The EVENT GAMES tile opens the ByBet Admin exactly as it is (passcode sign-in, then Spin the Wheel and Physical Game Tracker). Its passcode is the only gate, so anyone who can sign in to Event Hub and has the passcode can use it.
- Players use the public links the admin creates: `.../bybet/index.html?e=EVENTCODE` for the wheel and `.../bybet/tracker.html?g=GAMECODE` for staff. No Event Hub sign-in is needed.
- The Google Apps Script link lives in `bybet/config.js` and `bybet/index.html`.

## Notes
- A public GitHub Pages site is visible to anyone with the link. Firebase web keys are meant to be public; the Firestore rules protect your data.
- To update later, replace the changed files in the repo and commit. Pages redeploys in about a minute.
- Earlier README notes about Firebase setup, approvals, printing and inventory still apply to the HR Forms and Inventory modules.

## What was added (nothing was removed)
- **Event Games (ByBet Admin):** live events/games are listed first; an **Event Details** tile sits beside **Post Event Report**; a new **Event Data** tab (Spin the Wheel and Physical Games) shows Date, Event, Total FTD's, 21+ and Below 21 from the same Google Sheet data, with search, a below-21 filter, refresh and CSV download.
- **Assets/Links:** every Open/Copy made inside Event Hub is logged. Each link has **Who opened this link** (people, counts, copies, latest activity). Visible to HR and to the person who encoded the link. Opens made outside Event Hub cannot be recorded.
- **HR Forms:** your entries autosave as a draft on this device and come back if the page closes. The printable forms (OB, RD OT, Overtime, Change Schedule) only show your **saved** details; a banner appears while you have unsaved changes. New card **Overtime without Official Business** (date, time, required reason). It flows into the Overtime form marked "(No Official Business)".
- **Liquidation (new tile):** encode expenses (expense-type dropdown, auto totals, paste rows from a sheet), ByBet report preview matching the template, **Print**, **Copy for Google Sheet**, **Download CSV**, and saved liquidations in Firestore.
- **Google Sheet sync for Liquidation:** deploy `liquidation-sync.gs` in your Apps Script and paste its URL in `liquidation-config.js`.
- **Firestore rules:** republish `firestore.rules` (new collections `asset_views` and `liquidations`).

## Event Details and Promoter's Details (update)
- **Event Details** now lists: Event date, Event name (DATE-NAME-LOCATION, uppercase), Post event link, Event photo link, Event promoter's photo link, Promoter name and username, and Actual no. of participants. **Copy for message** copies all of it as lines ready to paste; each field also has its own Copy button.
- **Post Event Report** has a new **Promoter's Details** part (name, username, Old/New, optional photo link). It is NOT sent to the Google Sheet. It is saved on the device/browser used (same as report drafts) and shown in Event Details.
- A **Liquidation** button sits to the right of Post Event Report on every event and game. It opens the Liquidation tile with that event filled in (or the saved liquidation for it).

## Update: Event Games, Inventory (latest)
- **Liquidation** is no longer a sidebar tile. It lives inside Event Games (the Liquidation tab beside ByBet Admin, and the Liquidation button beside Post Event Report).
- **Spin the Wheel (public page):** when a prize runs out, the wheel refreshes, the prize is removed, and the spin continues automatically. The wheel also refreshes after each spin. The server (Apps Script) still decides the prize, so if it keeps returning an out-of-stock error, check the stock rules in the script.
- **New event:** has an Event location field. The event is saved as "Event Name - Event Location".
- **Post Event Report:** Promoter's photos + **Create promoter photo Google Doc** (needs `promoter-doc.gs` added to your Apps Script, then redeploy). The Doc link goes into Event Details. Promoter details stay out of the Google Sheet.
- **Event Details:** promoter type (Old/New) is on the same line as promoter name and username.
- **Inventory:** HR picks the stock out approver in HR Admin. Requesters see availability and send a request; the approver sets Box (unlimited) and Label (dropdown with "Add new label") and approves, which deducts the stock. Stock In category is a dropdown with "Add new category". The dashboard is split into Merchandise, Event Supplies and Event Collaterals with the In stock level integrated.
- **Firestore rules:** republish `firestore.rules` (new `stock_requests`, `config/inv`, and stock out is limited to the approver).

## Update: separate Liquidation, many promoters, Employee Access (V4)
- **Liquidation is its own tile** in the sidebar (LIQUIDATION), beside HR FORMS, INVENTORY, ASSETS/LINKS and EVENT GAMES. It is no longer a tab inside Event Games.
- **Event Liquidation (inside Event Games):** the Liquidation button beside Post Event Report on every event and game now opens a window on top of Event Games. Closing it (Close button or Esc) brings you back to the same event list. It is separate from the Liquidation tile: one per event, saved in Firestore `event_liquidations` and sent to the `EVENT LIQUIDATION` tab of your Google Sheet. Everyone approved can view it; the person who made it (or HR) can edit it.
- **Post Event Report > Promoter's Details:** you can add as many promoters as needed (Add another promoter / Remove). Each promoter has name, username, Old/New, photos and their own Google Doc. Event Details lists them numbered (names and photo links) and Copy for message includes all. Details saved with the old single-promoter format still show.
- **Employee Access (HR only):** HR Forms > HR Admin area > **Employee Access**. Tick HR Forms, Inventory, Assets/Links, Event Games, Liquidation for each employee, Save, or use Preview to see exactly what they see. Employees only see the tabs they are allowed. An employee HR has not set yet can use everything (so nobody is locked out after this update). HR always sees everything.
- **Deploy:** upload all files to GitHub (replace the old ones), publish `firestore.rules` in Firebase, and replace the Apps Script code with `liquidation-sync.gs` (plus `promoter-doc.gs`) then Deploy > Manage deployments > Edit > New version.
- **Note:** the hidden tabs are enforced in the app. Firestore rules still only check that a person is signed in and approved, so this controls what employees see, not a hard security wall per collection.
