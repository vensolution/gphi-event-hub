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
