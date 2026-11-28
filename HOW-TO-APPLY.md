# How to apply Phase 1 (step by step)

## 1. Back up
In GitHub, create a branch or copy of your repo. In Firebase > Firestore > Rules, copy the current rules into a text file.

## 2. Upload the files to GitHub
Replace these (same names, same folders):
- `index.html`
- `bybet/admin.html`
- `firestore.rules`
- `README.md`
Add this NEW file in the main folder, next to `index.html`:
- `ui-extras.js`
Commit. Wait 1 to 3 minutes. Open the site and press Ctrl+Shift+R (phone: clear the site data or open a private tab).

## 3. Publish the Firestore rules
Firebase console > Firestore Database > Rules > paste the new `firestore.rules` > Publish.
(Without this, Profile saves, the logo and profile requests will say "permission denied".)

## 4. Test (about 10 minutes)
1. Top bar: tap the moon or sun. The page switches between dark and light. Reload: it remembers.
2. Top bar shows the Philippine time. Compare it with a clock set to Manila.
3. Click PROFILE (top of the sidebar). Add a photo. Edit your phone and click Save. Request a change of your name.
4. Sign in as HR > HR Forms > Employee Access. Approve or reject the name request. Upload a new logo.
5. Open a password box and press Caps Lock. A yellow warning appears.
6. Event Games > ByBet Admin > Physical Game Tracker > Event Data. Each game shows its own numbers.
7. On a phone: buttons are bigger and typing does not zoom the page.

## If something is wrong
- Old screens: hard refresh (step 2).
- "permission denied": publish the rules (step 3).
- Roll back: put your backup files back in GitHub.
