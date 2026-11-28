# Phase 3: Leave Application + Messages

Changed files: `index.html` and `firestore.rules`. No Apps Script change.

## 1. Back up
Copy your GitHub repo (or make a branch). Copy your current Firestore rules into a text file.

## 2. Upload to GitHub (same folders)
- `index.html`
- `firestore.rules` (for your records)
Commit, wait 1 to 3 minutes, then Ctrl+Shift+R (phone: private tab).

## 3. Publish the rules (needed for Messages)
Firebase console > Firestore Database > Rules > paste the new `firestore.rules` > Publish.
Without this, Messages shows "permission denied". New: the `chats` collection.

## 4. Test Leave Application
1. HR Forms > Encode / Request > "Leave Application" card: Date hired, Type of leave, dates (days count by itself), Reason, Notes. Press Save.
2. HR Forms > Leave Application tab: two copies on one page. Print at 100%, margins none.
3. Encode / Request > Submit for approval > Leave Application. It goes Direct manager > Admin > HR. Check My Requests and the approver's Approvals.

## 5. Test Messages (use two accounts)
1. Sidebar > new MESSAGES tile (everyone can use it).
2. + New > tick one member > Start chat > send. The other account sees a red badge on the tile and a red dot in the list.
3. When the other person opens the chat, your message shows "Seen".
4. + New > tick two or more members, type a group name > Start chat.
5. HR account: open "Announcements" (pinned on top) and post. Everyone can read it, only HR can post.
6. On a phone the list and the chat show one at a time, with a back arrow.

## Known limits
- Messages are in-app only: no email or phone push notification. The badge shows while the site is open.
- Group members cannot be changed after the group is created.
- Messages cannot be edited or deleted yet.
- The Leave form layout is built from your text list. Send the paper form photo to match it exactly (logo, leave types with Paid / Unpaid boxes, columns).

## Roll back
Put your backup files back in GitHub and republish the old rules.
