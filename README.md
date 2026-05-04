# Nike Pacing Automation
### Why was this made?
- Allows us to keep track of all the currently running campaigns so that we can flag for any discrepancies in spending. I.e., whether a campaign is spending too fast or not enough.

- Only thing automated about our previous workflow was the data pulling in
- Each time a new campaign was launched, someone had to create a new row in the respective platform category, guess-and-check on the correct fields to search by, and often troubleshoot the excel formulas summarizing spend, impressions, etc.
- **Very error prone!!**
- Saves an estimated **2 hours per week!**

### What is it?
- This spreadsheet summarizes the spend of each campaign week-to-date (previous Sunday to the current day).
- Refreshes the performance and budget information every hour.
- Compares actualized spend to the spend goals listed in the Platform Level Details AirTable
- Automatically associates each spend goal record to a set of campaigns within our Alli Omnichannel dataset, and organizes it in an easily digestible view
- **No manual spreadsheet updates required!**

### Gotchas to lookout for
Sometimes, the automation has trouble associating a budget goal record in PLD to an actual campaign, and vice versa. This is because the *join keys* may not match. It looks for the following to associate campaigns to budgets:

- PLD Platform column matches the partner name of the campaign (DV360, Google Ads, etc…). Think: what DSP am I launching this campaign from.
- PLD publisher column matches the publisher in the campaign (YouTube, PadSquad, Spotify, Openx, etc…). Think: where will my customer see this ad?
- PLD “Initiative/Campaign” column appears in the campaigns’ initiative or sub campaign.
Something I’ve noticed while working on this automation is Nike has many different names for the same campaign. Ex: Soccer, NTK, World Cup, etc. The workflow gets very confused by these and will fail to link the campaigns to goals if they don’t have the same name. They will show as “NULL/UNATTRIBUTED”. Ensure campaigns in AirTable match campaign data you’re launching from

### Future ideas:
- Daily slack message
- Multiple time range tabs (WTD, MTD, Last Week, etc)
