# Windows Update Watchdog
A windows 10 & 11 update management tool. <br>
The newest update introduces the Watchdog tool. <br><br>
<img width="414" height="138" alt="image" src="https://github.com/user-attachments/assets/18abaefa-14c7-4a96-82d6-adb292581bd8" />


A python script with a HTML/CSS UI that sits in your system tray and checks periodically (10 seconds - or change to a value you want) whether "wuauserv" "UsoSvc" update services are running in the background, including Defender "WinDefend", "WdNisSvc", "Sense" services. If Update services are running, it will report a red light, and if defender is not running, it will flag a red light. You can ignore defender statuses (option on UI), and run the kill-script on the Update services. Incredibly tiny script with a small <50mb Memory fingerprint.


There is also the manual menu-driven PowerShell version, that will do the same: disable Windows OS updates, clear update caches while keeping Microsoft Defender security updates working.


![wmutps](https://github.com/user-attachments/assets/7c20c13d-fefb-489b-99f0-0674d7599389)

---

# What This Tool Does

- Disable Windows Update services, policies, and scheduled tasks
- Preventing automatic OS and feature updates
- Clear downloaded and pending Windows Update files
- Keep Microsoft Defender signature updates enabled
- Restore everything back to default & safe to re-run

![disable](https://github.com/user-attachments/assets/0616c32d-9cd5-42a5-969e-5ee0c677906c)

<b><h3>Instructions:</h3></b>
<b>Run as Administrator</b>

<u>Option 1</u>: Run the python tool and select "Run All" if light is Red (update is running). Services should detect disabled and light will turn green.<br>
<u>Option 2</u>: Run the PowerShell script and select 1 to run all. 

If update icon is stuck in tray, reboot system.

Using this tool Windows will still self-heal if Microsoft pushes OTA with persistence, but it will take time. Many hours, often a day or 2. If Updates persist, re-run the tool.
Reboot and it's cleared. 
<img width="1007" height="132" alt="image" src="https://github.com/user-attachments/assets/a33bd051-8e35-4a61-958d-1158f9a393a0" />



