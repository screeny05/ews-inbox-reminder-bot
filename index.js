const EWS = require('node-ews');

const debug = (...obj) => console.log(require('util').inspect(...obj, false, Infinity, true));

const itemShapeFactory = attributes => ({
    't:BaseShape': 'IdOnly',
    't:AdditionalProperties': {
        'FieldURI': attributes.map(attr => ({
            attributes: {
                FieldURI: attr
            }
        }))
    }
});

const calendarViewFactory = (start, end, maxEntries) => ({
    attributes: {
        MaxEntriesReturned: maxEntries,
        StartDate: start.toISOString(),
        EndDate: end.toISOString()
    }
});

class EwsClient {
    constructor(username, password, host){
        this.ews = new EWS({ username, password, host });
    }

    async _invoke(method, data){
        return this.ews.run(method, data, {
            't:RequestServerVersion': {
                attributes: {
                    Version: 'Exchange2016'
                }
            }
        });
    }

    async findCalendarItems(folderId, calendarView){
        const res = await this._invoke('FindItem', {
            'ItemShape': itemShapeFactory([]),
            'CalendarView': calendarView,
            'ParentFolderIds': {
                't:DistinguishedFolderId': {
                    attributes: {
                        Id: folderId
                    }
                }
            }
        });

        return res.ResponseMessages.FindItemResponseMessage.RootFolder.Items.CalendarItem;
    }

    async getCalendarItems(ids, shape = []){
        ids = ids.map(idObj => (idObj.ItemId))

        const res = await this._invoke('GetItem', {
            ItemShape: itemShapeFactory(shape),
            ItemIds: {
                't:ItemId': ids
            }
        });
        return res.ResponseMessages.GetItemResponseMessage.map(item => item.Items.CalendarItem);
    }

    async addInboxReminder(ids){
        const res = await this._invoke('UpdateItem', {
            attributes: {
                MessageDisposition: 'SaveOnly',
                ConflictResolution: 'AlwaysOverwrite',
                SendMeetingInvitationsOrCancellations: 'SendToNone'
            },
            ItemChanges: {
                't:ItemChange': ids.map(id => ({
                    't:ItemId': id.ItemId,
                    't:Updates': {
                        't:SetItemField': {
                            't:FieldURI': {
                                attributes: {
                                    FieldURI: 'calendar:InboxReminders'
                                }
                            },
                            't:CalendarItem': {
                                'InboxReminders': {
                                    'InboxReminder': [{
                                        ReminderOffset: 15,
                                        IsOrganizerReminder: false,
                                        OccurenceChange: 'None',
                                        IsImportedFromOLC: false,
                                        SendOption: 'User',
                                        Message: 'Beep-boop - i\'m a bot!'
                                    }]
                                }
                            }
                        }
                    }
                }))
            }
        });
    }
}


(async () => {
    const startDate = new Date();
    const endDate = new Date();
    endDate.setDate(endDate.getDate() + 30);

    const { EWS_USER, EWS_PASS, EWS_HOST } = process.env;
    if(!EWS_USER || !EWS_PASS || !EWS_HOST){
        throw new Error('Required env-params missing - EWS_USER, EWS_PASS, EWS_HOST');
    }

    const ews = new EwsClient(EWS_USER, EWS_PASS, EWS_HOST);
    const searchResults = await ews.findCalendarItems('calendar', calendarViewFactory(startDate, endDate, process.env.CHECK_ITEMS || 30));
    const items = await ews.getCalendarItems(searchResults, ['item:Subject', 'item:ReminderIsSet', 'calendar:InboxReminders', 'calendar:CalendarItemType']);

    // add reminder only for items which
    // * not have an inboxreminder already
    // * have a notification
    // * are not recurring events
    const addInboxReminderItems = items.filter(item => !item.InboxReminders && item.ReminderIsSet === 'true' && item.CalendarItemType !== 'Occurrence');

    console.log('adding reminder for:\n* ' + addInboxReminderItems.map(item => item.Subject).join('\n* '))

    await ews.addInboxReminder(addInboxReminderItems)
})();
