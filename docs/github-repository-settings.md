# GitHub repository settings (description, topics, security, branch protection)

These steps are done in the **GitHub web UI** (or with the **[GitHub CLI](https://cli.github.com/)** `gh` if you install it and `gh auth login`). This repo does not automate them in code.

**Repository:** [oyebiyisunday/product-costing-vendor-sourcing](https://github.com/oyebiyisunday/product-costing-vendor-sourcing)

---

## 1. Description and topics

### Web UI

1. Open the repository on GitHub.
2. Click **Settings** (top bar of the repo — you must be the **owner** or have **admin** access).
3. Under **General**, scroll to **Repository details**.
4. **Description** — suggested text (edit to taste):

   > Multi-vendor BOM costing in Excel: dynamic vendor pricing (XLOOKUP/SWITCH), batch scaling, and a Python/openpyxl generator. MIT licensed.

5. **Website** (optional) — e.g. a portfolio URL or leave blank.
6. **Topics** — click the **gear** next to “Topics” on the **Code** tab *or* add them where your layout shows topics. Typical tags for discoverability:

   `excel` · `bom` · `manufacturing` · `openpyxl` · `python` · `microsoft-365` · `costing` · `vendor-management`

### GitHub CLI (optional)

After `gh auth login`:

```bash
gh repo edit oyebiyisunday/product-costing-vendor-sourcing ^
  --description "Multi-vendor BOM costing in Excel: dynamic vendor pricing (XLOOKUP/SWITCH), batch scaling, and a Python/openpyxl generator. MIT licensed." ^
  --add-topic excel --add-topic bom --add-topic manufacturing --add-topic openpyxl --add-topic python --add-topic costing
```

*(On macOS/Linux, use `\` line continuation instead of `^`.)*

---

## 2. Private vulnerability reporting

This lets researchers report **security issues** without opening a public issue (aligned with [`SECURITY.md`](../SECURITY.md)).

### Web UI

1. Repo **Settings** → **Code security and analysis** (under “Security” in the left sidebar; exact label may be **Security** → **Advanced Security** depending on GitHub’s layout).
2. Find **Private vulnerability reporting**.
3. Click **Enable** (or **Start enablement** if GitHub shows an onboarding step).

If you do not see the option: your account/org plan or permissions may restrict it; confirm you are an **admin** on the repository.

### Verify

- Repo **Security** tab → **Reporting** / advisories area should reflect that private reporting is available, per [GitHub Docs — configuring private vulnerability reporting](https://docs.github.com/en/code-security/security-advisories/working-with-repository-security-advisories/configuring-private-vulnerability-reporting-for-a-repository).

### GitHub CLI / API (optional)

If your token has `repo` or `admin:org` as required by GitHub’s API for this feature, you can use the REST API; many users find the **web UI** simpler. See GitHub’s REST reference for **repository security analysis** updates if you need automation.

---

## 3. Branch protection for `main` (optional)

Goal: **no direct pushes to `main`** (optional) and/or **required CI** before merge.

### Web UI (classic branch protection rule)

1. **Settings** → **Branches**.
2. **Branch protection rules** → **Add branch protection rule**.
3. **Branch name pattern:** `main`.
4. Recommended toggles (pick what fits a **solo** vs **team** repo):

   | Rule | Solo maintainer | Team |
   |------|-----------------|------|
   | **Require a pull request before merging** | Optional (you may still want PRs for history) | Recommended |
   | **Require status checks to pass before merging** | Recommended | Recommended |
   | **Require conversation resolution before merging** | Optional | Often useful |
   | **Do not allow bypassing the above settings** | Often **off** so you can hotfix | **On** for production repos |
   | **Include administrators** (enforce on admins) | Your choice | Often **on** |

5. **Status checks** — enable **Require status checks to pass before merging**, then search for the CI job. This repository’s workflow is [`.github/workflows/ci.yml`](../.github/workflows/ci.yml); the job id is **`verify`**. In the UI you may see:

   - `verify`, or  
   - `CI / verify`  

   Select the check that runs on **pull requests** to `main`.

6. Save the rule.

### Rulesets (newer alternative)

GitHub also offers **Repository rules** / **Rulesets** (Settings → **Rules** → **Rulesets**). You can require the same checks there; use GitHub’s docs for **rulesets** if you prefer that model over classic branch rules.

### After protection is on

- Push only via **PR**, or temporarily allow bypass for admins if you need a direct push.
- First PR after adding checks: confirm the **CI** workflow is green so the required check appears in the branch protection dropdown.

---

## 4. Related settings worth enabling

| Setting | Where | Why |
|--------|--------|-----|
| **Issues** | Settings → **General** → Features | Bug reports and CoC enforcement. |
| **Dependabot alerts** | Settings → **Code security and analysis** | Alerts for vulnerable dependencies (works with [`.github/dependabot.yml`](../.github/dependabot.yml)). |
| **Default branch** | Settings → **Branches** | Should be **`main`**. |

---

## Quick reference links

- [Repository settings](https://github.com/oyebiyisunday/product-costing-vendor-sourcing/settings)  
- [Branches / protection](https://github.com/oyebiyisunday/product-costing-vendor-sourcing/settings/branches)  
- [Code security and analysis](https://github.com/oyebiyisunday/product-costing-vendor-sourcing/settings/security_analysis)  
- [Actions](https://github.com/oyebiyisunday/product-costing-vendor-sourcing/actions)
