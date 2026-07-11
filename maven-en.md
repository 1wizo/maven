# Maven — A Straightforward Guide to Office LTSC + Activation

Let's be honest about what this guide is: a no-nonsense walkthrough for getting Microsoft Office LTSC installed on Windows without paying for a subscription, and then activating both Windows and Office so they stop nagging you. It uses tools Microsoft themselves publish, plus an open-source project called MAS that handles the activation part. No sketchy ISOs, no cracks downloaded from random forums, no "activation tools" that are actually malware.

I've also included a section for Linux users at the end, because not everyone runs Windows and the Office situation over there is genuinely confusing.

---

## What You'll Need

Before we start, make sure you've got:

- **Windows 10 or 11** — any edition works, including LTSC. Windows 7/8.1 technically work for activation but not for Office LTSC 2021+.
- **About 5 GB of free space** — Office itself is 2–4 GB, plus room for the installer and temp files.
- **A stable internet connection** — you're downloading Office directly from Microsoft's CDN, so the speed depends on your pipe.
- **Administrator access** — you'll need this for both the Office installer and the activation script.
- **PowerShell 5.1 or later** — already installed on every Windows 10/11 machine, so you probably don't need to worry about this.

Total time commitment: about 15–30 minutes, mostly waiting for the Office download to finish.

---

## Part 1 — Installing Office LTSC

The approach here might feel a bit roundabout, but it's actually the cleanest way to do it. Instead of downloading some cracked Office ISO, we're going to use Microsoft's own Office Deployment Tool (ODT) — the same tool enterprise IT departments use to deploy Office across hundreds of machines. The difference is we're pointing it at the Volume License channel, which gives us the LTSC build that MAS can activate.

### Step 1 — Grab the Office Deployment Tool

Head over to Microsoft's official download page:

```
https://www.microsoft.com/download/details.aspx?id=49117
```

Click the **Download** button, run the `officedeploymenttool_*.exe` file that comes down, accept the license agreement, and extract everything to a folder you'll remember — something like `C:\OfficeInstall\` works well.

What you'll end up with is `setup.exe` (this is the important one — keep it) and a handful of sample XML files you can safely ignore. We're going to generate our own XML in the next step.

### Step 2 — Build Your Configuration

This is the part where you tell the installer exactly what you want. Microsoft has a web-based tool for this, and it's genuinely well-made:

```
https://config.office.com/deploymentsettings
```

Sign in with any Microsoft account — a free Outlook.com account works fine, you don't need anything special. Once you're in, you'll see a wizard. Here's what to set:

**Under "Office Products":**
- **Products:** Choose "Office LTSC Professional Plus 2021 - Volume License" (or 2024 if you want the newer version)
- **Update channel:** This is critical — set it to "Office LTSC 2021 Perpetual Enterprise." If you pick the wrong channel, you'll get the consumer build instead of the Volume License build, and MAS won't be able to activate it.
- **Apps:** Tick the ones you actually want. Word, Excel, PowerPoint, Outlook, and OneNote are the core five. Visio and Project are available too if you need them. Untick the junk — Clipchamp, Bing, and whatever else Microsoft is bundling these days.

**Under "Language":**
Pick whatever you want. English (US) is the safe default, but French, Arabic, Spanish, German, and dozens of others are all there.

**Under "Settings":**
- Installation platform: **64-bit** (there's no good reason to use 32-bit in 2025)
- Update automatically: **Yes**
- Pin to taskbar: **No** (nobody wants that)
- Remove older versions: **Yes** (this cleanly uninstalls any existing Office before installing LTSC)

**Under "Licensing and Activation":**
- Accept EULA: **Yes**
- Product key: **leave blank** (the Volume License build doesn't need one embedded in the config)
- AUTOACTIVATE: **FALSE** — we'll handle activation separately with MAS

Once everything looks right, click **Export** in the top right, choose **XML** as the format, and save the file as `configuration.xml` in the same folder as `setup.exe`.

### Step 3 — Run the Installer

Open Command Prompt as administrator (Start menu → type `cmd` → right-click → Run as administrator), navigate to your Office folder, and run the installer:

```batch
cd /d C:\OfficeInstall
setup.exe /configure configuration.xml
```

A small Office window will pop up and start downloading. This is the part where you go make coffee — it's pulling 2–4 GB directly from Microsoft's CDN, so depending on your connection, you're looking at anywhere from 5 to 20 minutes. The install runs silently in the background; you don't need to click through any wizards.

When it's done, you'll see Word, Excel, PowerPoint, and the rest in your Start menu. They'll launch fine, but you'll notice an "Activate Office" banner — that's expected, and we'll deal with it in Part 3.

### Step 4 — Quick Verification

Open Word. It should launch without asking for a product key, and you should be able to create and save documents. The activation banner will be there, but the app itself is fully functional. If Word won't launch at all, something went wrong with the install — check that `setup.exe` exited with code 0 in your command prompt.

<details>
<summary>Want to write the XML by hand instead? Here are the product IDs</summary>

```xml
<Configuration>
  <Add OfficeClientEdition="64-bit" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume">
      <Language ID="en-us" />
    </Product>
    <Product ID="Visio2021Volume">
      <Language ID="en-us" />
    </Product>
    <Product ID="Project2021Volume">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="FALSE" />
</Configuration>
```

| Product ID | What it installs |
|---|---|
| `ProPlus2021Volume` | Office LTSC Pro Plus 2021 |
| `Standard2021Volume` | Office LTSC Standard 2021 |
| `Visio2021Volume` | Visio LTSC 2021 |
| `Project2021Volume` | Project LTSC 2021 |
| `ProPlus2024Volume` | Office LTSC Pro Plus 2024 |
| `Visio2024Volume` | Visio LTSC 2024 |
| `Project2024Volume` | Project LTSC 2024 |

If you want the 2024 versions, just swap `Channel="PerpetualVL2021"` to `"PerpetualVL2024"` and use the 2024 product IDs.
</details>

---

## Part 2 — Activating Windows

Now let's talk about Windows activation. The tool we're using is called MAS (Microsoft Activation Scripts). It's an open-source project that's been around for years, is actively maintained on GitHub, and works by generating a legitimate hardware-bound digital license — the same kind you'd get if you bought Windows from Microsoft and logged into your Microsoft account.

The method we'll use is called HWID. It ties a permanent digital license to your hardware, so it survives reinstalls, feature updates, and even most hardware changes. This is the same activation type Microsoft gives to people who upgrade from Windows 7/8 for free.

### Step 1 — Open PowerShell as Administrator

Press the Windows key, type `powershell`, right-click **Windows PowerShell**, and select **Run as administrator**. Click **Yes** on the UAC prompt.

### Step 2 — Run the MAS Launcher

Paste this single line and press Enter:

```powershell
irm https://get.activated.win | iex
```

If `get.activated.win` is blocked by your DNS or firewall (some corporate networks do this), use the mirror instead:

```powershell
irm https://massgrave.dev/get | iex
```

A text-based menu will appear. It looks like a DOS program from 1995, but it works perfectly.

### Step 3 — Choose Your Activation Method

Type `1` and press Enter to select "Activate Windows." You'll then see three methods:

1. **HWID** — This is the one you want. It creates a permanent digital license tied to your hardware. Best for Windows 10 and 11. Use this.
2. **KMS38** — Valid until the year 2038. Use this if HWID fails for some reason, or if you're on Windows Server.
3. **Online KMS** — Renews every 180 days. Only use this if you're on Windows 7 or 8.1, which don't support HWID.

Pick option 1 (HWID), confirm with `Y`, and wait. It usually takes 10–30 seconds. You'll see a message saying "Activation successful" when it's done.

### Step 4 — Verify

Type `5` in the MAS menu and press Enter to check activation status. It should say "Windows is permanently activated."

That's it. The license is tied to your motherboard and CPU signature, so it'll survive Windows updates, feature upgrades (like 23H2 → 24H2), and even a clean reinstall on the same hardware.

---

## Part 3 — Activating Office

Office is installed but not activated yet. We're going to use a different MAS method called Ohook for this one. Ohook works by patching Office's activation check at the binary level — it's permanent, works offline, and survives Office updates. It's the most reliable Office activation method MAS offers.

### Step 1 — Close All Office Apps

This is important. Make sure Word, Excel, PowerPoint, Outlook, Visio, and Project are all completely closed. Check your system tray (bottom-right corner of the taskbar) and quit any Office icons hiding there. Ohook needs to modify files that Office locks while running, so if anything is open, the activation will fail.

### Step 2 — Run MAS Again

Same launcher as before:

```powershell
irm https://get.activated.win | iex
```

### Step 3 — Activate

In the MAS menu:
- Type `2` and press Enter (Activate Office)
- Type `1` and press Enter (Ohook)
- Type `Y` to confirm

Wait for "Activation successful." This usually takes 5–15 seconds.

### Step 4 — Verify

Type `5` in the MAS menu to check status. All your Office products should show "LICENSED." Open Word, go to File → Account, and you'll see "Microsoft Office LTSC Professional Plus 2021" with no activation banner. You're done.

---

## Part 4 — Linux Office Alternatives

If you're reading this on Linux, the situation is different. Microsoft Office doesn't run natively on Linux, and honestly, it probably never will. But you've got two solid options depending on what you actually need.

**OnlyOffice** is a native Linux office suite that's free, open-source, and reads/writes `.docx` files better than anything else I've tried. If you just need to write documents, build spreadsheets, and make presentations — and you want them to look right when your Windows-using colleagues open them — this is the one.

**LinOffice** is something else entirely. It's a 1-click installer that spins up a Windows VM in a container and runs the actual Microsoft Office 2024 inside it, then presents each Office app as if it were a native Linux window. It's heavier (needs 8 GB RAM and KVM), and you still need an Office license, but you get 100% perfect MS Office compatibility with full VBA macro support.

### Which one should you actually use?

Honestly, for 90% of Linux users, **OnlyOffice** is the right answer. It's light, it's free, it doesn't require a VM, and the `.docx` fidelity is excellent. The only reason to go with LinOffice is if you absolutely need real Microsoft Office — maybe your job requires VBA macros that OnlyOffice can't run, or you need perfect rendering of complex Excel models, or you need Outlook for Exchange integration.

| | OnlyOffice | LinOffice |
|---|---|---|
| **What it is** | Native Linux office suite | Real MS Office 2024 in a Windows VM |
| **License** | AGPL v3 (free, open-source) | AGPL v3 (free, open-source) — forked from WinApps |
| **MS Office fidelity** | Excellent — native OOXML | Perfect — it IS Microsoft Office |
| **Includes** | Word, Excel, PowerPoint equivalents | Real Word, Excel, PowerPoint, OneNote, Outlook |
| **Resource usage** | Light — about 450 MB | Heavy — 8 GB RAM, 64 GB disk, full Windows VM |
| **Hardware** | Any (x86_64, ARM, etc.) | x86_64 only + KVM virtualization |
| **Activation needed** | No — completely free | Yes — Office 2024 key or M365 subscription (MAS works too) |
| **Linux integration** | Native — file associations, notifications, all good | Good via FreeRDP; some quirks on Wayland/multi-monitor |

### Installing OnlyOffice

Pick whichever method matches your distro:

```bash
# Snap (works on any distro that has snapd)
sudo snap install onlyoffice-desktopeditors

# Flatpak (works on any distro that has Flatpak)
flatpak install flathub org.onlyoffice.desktopeditors

# Ubuntu / Debian (.deb — download from onlyoffice.com/download-desktop.aspx)
sudo apt install ./onlyoffice-desktopeditors-amd64.deb

# Arch (AUR)
yay -S onlyoffice-bin
```

### Installing LinOffice

LinOffice has a quickstart script that handles everything automatically — it detects your distro, installs all dependencies, downloads the installer, and launches the GUI:

```bash
curl -sSL https://github.com/eylenburg/linoffice/raw/refs/heads/main/quickstart.sh -o quickstart.sh \
  && chmod +x quickstart.sh && ./quickstart.sh
```

If you prefer to install dependencies manually:

```bash
# Ubuntu 24.04+ / Debian 13+
sudo apt install podman podman-compose freerdp3 python3 \
  python3-pyside6 python3-pyside6.qtwidgets python3-pyside6.qtuitools

# Fedora
sudo dnf install podman podman-compose freerdp python3 python-pyside6

# Arch
sudo pacman -Syu podman podman-compose freerdp python pyside6

# openSUSE
sudo zypper install podman podman-compose freerdp python3 python3-pyside6
```

Then download the repo, extract it, and run either `./setup.sh` (CLI) or `python3 gui/linoffice.py` (GUI).

**Heads up on hardware requirements:** LinOffice needs an x86_64 CPU with KVM support (check your BIOS), at least 8 GB RAM, and about 64 GB of free disk space. The initial download from Microsoft is around 8 GB, and the whole install takes about 15 minutes on a decent connection. Also, avoid FreeRDP version 3.23 — it has a known bug. Use 3.22 or wait for 3.24.

Once installed, you get a `linoffice` command:

```bash
linoffice word           # launch Word
linoffice excel          # launch Excel
linoffice powerpoint     # launch PowerPoint
linoffice onenote        # launch OneNote
linoffice outlook        # launch Outlook
linoffice windows        # full Windows desktop via RDP
linoffice update         # update Windows + Office
linoffice stopcontainer  # stop the VM (keeps your data)
```

The installer also creates `.desktop` entries, so Word, Excel, etc. show up in your app launcher like normal Linux apps. Files in `/home` are accessible from inside the VM, so you can open and save documents to your regular Linux home directory.

### What about Wine/Bottles/CrossOver?

You can technically run MS Office through Wine or Bottles, but honestly, it's fragile. Every Office update has a chance of breaking something, macro support is spotty, and activation gets weird. CrossOver is the most polished option but it's paid. LinOffice is the cleanest free solution if you really need real Office — it's essentially WinApps (its upstream project) with a bunch of automation layered on top.

---

## Quick Cheat Sheet

If you've done this before and just need the commands:

**Install Office:**
```batch
cd /d C:\OfficeInstall
setup.exe /configure configuration.xml
```

**Activate Windows:**
```powershell
irm https://get.activated.win | iex
```
→ `1` → `1` (HWID) → `Y`

**Activate Office:**
```powershell
irm https://get.activated.win | iex
```
→ `2` → `1` (Ohook) → `Y`

**Both at once (silent):**
```powershell
& ([ScriptBlock]::Create((irm https://get.activated.win))) -HWID -Ohook -S
```

**Check activation status:**
```batch
slmgr /xpr                    :: Windows
cscript ospp.vbs /dstatus     :: Office (run from the Office folder)
```

---

## Troubleshooting

Real problems I've actually seen, and how to fix them:

**Office installer errors out** — Make sure `configuration.xml` is in the same folder as `setup.exe`. If the XML is malformed, re-export it from config.office.com. The most common cause is a typo in the product ID or channel name.

**Install hangs at "Downloading"** — Usually a network issue. Disable any VPN or proxy, make sure `*.officecdn.microsoft.com` isn't blocked by your firewall, and check that you actually have 5+ GB free on your system drive.

**You got Office 365 instead of LTSC** — You picked the wrong update channel in the config. Go back to config.office.com and set it to "Office LTSC 2021 Perpetual Enterprise" — not "Current Channel" or "Monthly Enterprise."

**Visio or Project didn't install** — They need to be added as separate `<Product>` entries in the XML, not just ticked in the apps list. See the XML example in Part 1.

**"Can't install 64-bit Office with 32-bit already installed"** — Set "Remove older versions = Yes" in the config, or manually uninstall the old 32-bit Office from Settings → Apps first.

**MAS can't resolve `get.activated.win`** — Your DNS is blocking it. Use the mirror: `irm https://massgrave.dev/get | iex`

**Windows Defender deletes the MAS script** — This is expected. MAS triggers false positives in most antivirus software. Temporarily disable real-time protection, add the folder to exclusions, run MAS, then re-enable Defender. It's safe — MAS is open-source and audited.

**HWID fails with error 0xC004C003** — Try KMS38 instead (option 2). It's valid until 2038 and works on the same Windows builds. You can also run the built-in troubleshooter: MAS menu → `[6] Troubleshoot` → `Fix Licensing`, then retry.

**MAS doesn't detect Office** — You installed the consumer (Click-to-Run) build instead of the Volume License build. The fix is to rebuild your config with the `PerpetualVL2021` channel and reinstall.

**Activation says "180 days" instead of permanent** — You got Online KMS instead of HWID/Ohook. Re-run MAS and specifically pick HWID for Windows (option 1) and Ohook for Office (option 1).

**Activation disappeared after a hardware upgrade** — HWID licenses are tied to your motherboard and CPU. If you replace both, you may need to re-run HWID. It'll re-bind to the new hardware automatically.

**Network errors (0x80072EE2, 0x80072EFD)** — These are connectivity issues. Turn off VPN/proxy, allow `*.windows.com` and `*.microsoft.com` through your firewall, and try a different DNS resolver (1.1.1.1 or 8.8.8.8 often work better than ISP DNS).

---

## FAQ

**Can I do this on a fresh Windows install?**
Yes. The order doesn't matter — you can activate Windows first and install Office after, or vice versa. Both work fine on a clean Windows 10/11 with no additional setup.

**Do I need a Microsoft account?**
Only to sign into config.office.com and build the configuration XML. The Office installer itself doesn't require a login — it downloads directly from Microsoft's CDN with no authentication. You can use a throwaway Outlook.com account for the config tool.

**Will Office update itself?**
Yes. LTSC gets security patches through Windows Update. The "LTSC" part means "Long-Term Servicing Channel" — you get security fixes but no feature updates. The UI won't change every month like it does with Office 365. That's the whole point.

**Does activation survive Windows feature updates (like 23H2 → 24H2)?**
Yes. HWID creates a hardware-bound digital license that persists through feature upgrades. Ohook patches Office's activation files, which also survive updates. You shouldn't need to re-activate after any normal Windows update.

**Can I install just Visio and Project without Word/Excel?**
Yes. In config.office.com, untick everything except Visio and Project in the Apps section. Or if you're handwriting the XML, just include `Visio2021Volume` and `Project2021Volume` as the only `<Product>` entries.

**How do I switch from Office 365 to LTSC?**
Set "Remove older versions = Yes" in the config (or add `<RemoveMSI />` to the XML). When you run `setup.exe /configure`, it uninstalls 365 and installs LTSC in one pass. Your documents won't be touched.

**Does this work on Windows 7 or 8.1?**
MAS works on Windows 7/8.1 for Windows activation, but Office LTSC 2021/2024 requires Windows 10 or later. If you're on an old Windows version, your best bet is Office 2016 + Online KMS activation.

**Where are the MAS logs?**
`C:\ProgramData\MAS\` — this folder contains activation logs, diagnostic info, and ticket files. Check here if anything goes wrong and you want to debug.

**How do I uninstall Office later?**
Settings → Apps → Microsoft Office → Uninstall. Or you can use the ODT with an uninstall XML — see Microsoft's documentation for the schema.

**Can I activate Office 365 Family/Personal with MAS?**
Ohook works on consumer Office 365 builds. KMS doesn't — consumer 365 builds aren't Volume License, so KMS38 and Online KMS won't activate them. If you have a consumer 365 subscription, Ohook is your only MAS option.

**Is the `irm | iex` command safe to copy/paste?**
Yes. `get.activated.win` is the official shortlink maintained by the MAS project. It redirects to the script hosted in the `massgravel/Microsoft-Activation-Scripts` GitHub repository. Always verify you're using the official URL — there are impersonators out there.

**How do I update MAS to a newer version?**
Just run the same `irm https://get.activated.win | iex` command again. New versions automatically replace old ones. There's nothing to uninstall — MAS detects and overwrites the previous version.

---

## Useful Links

- **MAS project page:** https://massgrave.dev
- **MAS GitHub repository:** https://github.com/massgravel/Microsoft-Activation-Scripts
- **Latest MAS release (.zip):** https://github.com/massgravel/Microsoft-Activation-Scripts/releases/latest
- **Office Deployment Tool:** https://www.microsoft.com/download/details.aspx?id=49117
- **Office Customization Tool:** https://config.office.com/deploymentsettings

---

*This guide is for educational purposes. Always pull scripts from the official GitHub repository to ensure authenticity. The author and MAS project are not affiliated with Microsoft.*
