# AME Engineering Latur - MHT-CET College Predictor

A beautiful, high-performance, and responsive glassmorphic web application built to predict MHT-CET engineering college cutoffs and choices dynamically.

## 🚀 Core Features
1. **Premium Glassmorphic Dashboard**: Organically breathing backdrop glows, rich typography, fluid micro-animations, and gold accent indicators.
2. **Mutual Rank & Percentile Auto-Fetching**: Entering a merit rank instantly looks up the historical percentile (and vice-versa) in real-time, utilizing a database of 32,267 unique records.
3. **Dynamic Range Filtering Engine**: Automatically expands search bounds starting at `1500` steps and expanding by `1000` until at least 30 matches are populated.
4. **Curated Recommendations**: Recommends up to 5 DREAM colleges at the top, and lists target and safe colleges with a maximum cap of 45 options.
5. **Personalized Excel Downloads**: Tailored spreadsheet reports with student-specific filenames.

---

## 🌐 Deploying to GitHub Pages (Free Hosting)

GitHub Pages is the easiest way to launch this website for free! Follow these simple steps:

### Option A: Uploading via the GitHub Website (No Git Command Line Required)
1. **Create a GitHub Account**: If you don't have one, sign up at [github.com](https://github.com).
2. **Create a New Repository**:
   * Click **New** (or go to `github.com/new`).
   * Name your repository (e.g., `mht-cet-predictor`).
   * Keep it **Public** (required for free GitHub Pages hosting).
   * Do not check "Add a README file" (we already have one in the ZIP).
   * Click **Create repository**.
3. **Upload the Files**:
   * On the repository setup page, click the link that says **"uploading an existing file"**.
   * Unzip the generated `ame-engineering-latur-website.zip` on your computer.
   * Drag and drop the **7 files** (do *not* upload the ZIP itself, upload the files *inside* it):
     - `index.html`
     - `style.css`
     - `app.js`
     - `logo.png`
     - `colleges_data.json`
     - `rank_percentile_mapping.json`
     - `README.md`
   * Wait for the files to upload, then click **Commit changes** at the bottom.
4. **Enable GitHub Pages**:
   * Go to the **Settings** tab of your repository.
   * On the left sidebar, click **Pages** under the "Code and automation" section.
   * Under **Build and deployment**, change the Source branch from **None** to **main** (or `master`).
   * Click **Save**.
   * Wait 1–2 minutes, and refresh the page. A banner will appear at the top showing your live URL! (e.g., `https://username.github.io/mht-cet-predictor/`).

---

### Option B: Uploading via Git Command Line
If you are comfortable using Git in your terminal, run the following commands in the unzipped folder:
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/your-username/your-repo-name.git
git push -u origin main
```
Then, follow **Step 4** above to activate GitHub Pages.

---

## 📞 Counselling Helpline
For expert engineering admission support and counselling, reach out at **`+91 94226 11661`**.
