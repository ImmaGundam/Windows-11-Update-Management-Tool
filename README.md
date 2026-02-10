# Windows 11 Update Management Tool

A menu-driven PowerShell utility to disable Windows OS updates, clear update caches while keeping Microsoft Defender security updates working.

![wmutps](https://github.com/user-attachments/assets/7c20c13d-fefb-489b-99f0-0674d7599389)

---

# What This Tool Does

- Disable Windows Update services, policies, and scheduled tasks
- Preventing automatic OS and feature updates
- Clear downloaded and pending Windows Update files
- Keep Microsoft Defender signature updates enabled
- Restore everything back to default & safe to re-run

Instructions:
Using this script currently Windows will still self-heal, but it will take time. Many hours, often a day or 2. If Updates persist, run the script as admin and run 1 (run all). Reboot and it's cleared. 

WIP/To do: Watchdog app for monitoring and auto-running script

