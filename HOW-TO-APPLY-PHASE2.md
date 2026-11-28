# Phase 2 (part 1): Inventory upgrades + one Google Doc for all promoters

## 1. Back up
Copy your GitHub repo (or make a branch). Copy your current Firestore rules and your Apps Script into text files.

## 2. Upload to GitHub (same names, same folders)
- `index.html`
- `bybet/admin.html`
- `firestore.rules`
Commit, wait 1 to 3 minutes, then Ctrl+Shift+R (phone: private tab).

## 3. Publish the Firestore rules
Firebase console > Firestore Database > Rules > paste the new `firestore.rules` > Publish.
New: several approvers (`approverUids`) and the `stock_returns` collection. Without this you get "permission denied".

## 4. Replace the whole Apps Script (one file)
1. Open your Apps Script project (Code.gs). Ctrl+A, Delete.
2. Open `apps-script-COMPLETE-Post-Event-Report.gs` and paste the whole file.
3. Check Project Settings > Script properties: ADMIN_PASSCODE (and GEMINI_API_KEY if you use Improve writing) are still there.
4. Ctrl+S > Deploy > Manage deployments > pencil > Version: New version > Deploy. The URL stays the same.

## 5. Test
Inventory (use two accounts, an employee and the approver):
1. HR Forms > HR Admin > "Inventory: who approves stock out": add two people. Remove one with the x.
2. Employee: Inventory > Stock Out. Tap a category tile. The item list shows only that category. Fill Location and send.
3. Employee: tap the request under "My stock out requests". It expands.
4. Approver: a pop-up appears when the app opens or when a new request arrives. Approve it (Box + Label).
5. Employee: open the approved request > "Event cancelled? Return items". Choose the reason, Return to stock or Move to another event, set the condition of each item, Send.
6. Approver: Stock Out > "Returns waiting" > Confirm received. Items marked Good go back to stock. Damaged items do not.

Promoters:
1. Event Games > ByBet Admin > Post Event Report > Promoter's Details. Add 2 promoters with photos. Use the arrows to reorder.
2. Click "Create the Google Doc" in the blue box. One Doc opens with one section per promoter, 4 photos per row.

## Roll back
Put your backup files back in GitHub and republish the old rules / script version.
