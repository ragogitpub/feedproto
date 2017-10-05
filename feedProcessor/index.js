const $SP = require('sharepointplus');

module.exports = function (context, myQueueItem) {
    context.log('JavaScript queue trigger function processed work item', myQueueItem);

    var agencyName = myQueueItem.agencyName;
    var listName = myQueueItem.listName;
    var idField = myQueueItem.idField;

    var itemJSON = myQueueItem;
    itemJSON.PartitionKey = agencyName + '-' + listName;
    itemJSON.RowKey = myQueueItem[ idField ] + '-' + (new Date()).toISOString();
    context.bindings.tableContent = [ itemJSON ];

    context.done();
};
