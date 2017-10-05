const $SP = require('sharepointplus');

module.exports = function (context, myQueueItem) {
    context.log('JavaScript queue trigger function processed work item', myQueueItem);

    var agencyName = myQueueItem.nc4__agencyName;
    var listName = myQueueItem.nc4__listName;
    var idField = myQueueItem.nc4__idField;

    var itemJSON = myQueueItem;
    itemJSON.PartitionKey = agencyName + '-' + listName;
    itemJSON.RowKey = myQueueItem[ idField ] + '-' + (new Date()).toISOString();
    context.bindings.tableContent = [ itemJSON ];

    context.done();
};
