module.exports = function (context, myQueueItem) {
    context.log('JavaScript queue trigger function processed work item', myQueueItem);

    var itemJSON = myQueueItem;
    itemJSON.PartitionKey = 'Test';
    itemJSON.RowKey = itemJSON.CADItemNo + '-' + (new Date()).toISOString();
    context.bindings.tableContent = [ itemJSON ];

    context.done();
};
