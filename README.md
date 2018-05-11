# ews-inbox-reminder-bot

This project is born out of frustration with exchange.
Currently, there is no way to enable mail-reminders for all calendar-items.

So instead of enabling them by hand, this script uses ews to add the somewhat undocumented `calendar:InboxReminders` to all of your calendar entries which:

* do not already have an inboxreminder set
* have a reminder set
* are not recurring events
* happen between `now` and `now + 30 days`

This script is intended to run on heroku via the Scheduler add-on.

Configuration needed is done via environment-variables. Take a look at `.env`.
