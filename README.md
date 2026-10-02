# Kuwait Hospital – Meeting Room Booking

Web pages hosted on **GitHub Pages**; data, emails and sign-in handled by **Google Apps Script + Google Sheet**.

| Page | File | Address (after publishing) |
|---|---|---|
| Booking form | `index.html` | `https://<user>.github.io/<repo>/` |
| Coordinator portal | `coordinator.html` | `https://<user>.github.io/<repo>/coordinator.html` |
| Room tablet screen | `room.html` | `https://<user>.github.io/<repo>/room.html?room=Board%20Room` |

## 1. Backend (Google Apps Script)
1. Create a Google Sheet → **Extensions → Apps Script**.
2. Add two script files from the `apps-script` folder: **Code** and **Logo** (no HTML files needed).
3. Run **setup** once and allow the permissions.
4. **Deploy → New deployment → Web app** – *Execute as: Me*, *Who has access: Anyone*. Copy the `/exec` link.
5. In Code.gs set `WEB_APP_URL` (the `/exec` link) and `SITE_URL` (your GitHub Pages address), save,
   then **Manage deployments → Edit → New version → Deploy**.

## 2. Website (GitHub Pages)
1. Put the `/exec` link in **`js/config.js`**.
2. Upload everything to a GitHub repository.
3. **Settings → Pages → Deploy from a branch → main / (root) → Save.** The site is live in a minute or two.

## 3. Accounts and rooms
* Edit rooms in the **Rooms** tab of the Sheet.
* Sheet menu **Meeting Rooms → Add / update coordinator** – use `ALL` for a supervisor.
* **Meeting Rooms → Show links** lists the form, portal and one tablet link per room.

No passwords or keys are stored in this repository – passwords are hashed inside the Apps Script project.
