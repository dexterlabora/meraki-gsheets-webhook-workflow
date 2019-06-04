# meraki-gsheets-webhook-workflow
A Google Sheets script to capture Meraki webhooks. It then adds a workflow to reboot devices when a VPN connection has failed for some timespan.



# How it Works
A new Sheet tab will be created for each webhook alert type (i.e. VPN connectivity changed)
A new Row will be inserted into the sheet upon receiving events

# Setup
Publish Code (Tools -> Script Editor): Deploy as Web App. 
Execute App as you (authorize when prompted)
Access: Anyone, even Anonymous
Configure Meraki to point the webhook alerts to the URL provided
Configure Macro Trigger (Tools -> Script Editor): Current Projects Triggers --> Add Trigger :
- Function: macroRebootDevicesVPN
- Event Source: Time Driven
- Type of Time: Minutes timer
- Interval: Every minute

# Settings
Update the settings sheet with your API key, the base URL for your Meraki Dashboard and the timespan for running the workflow script


# Workflows
## VPN Reboot
- When a VPN connectivity changed event is received when the connectivity is false, the device will be staged for a reboot. 
- The devices to be rebooted will appear in the Queue: VPN-Reboot sheet.
- When the macro runs, it will check the occurredAt time for the alert and compare it against the timespan set in the settings sheet.
- If the time has been reached, the devices will be rebooted and the event will be recorded in the Logs: VPN workflow sheet
- If a new alert is received that states the connectiviy is true, the device will be removed from the reboot queue.
