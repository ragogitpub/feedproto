const $SP = require('sharepointplus');

module.exports = function (context, myQueueItem) {
    context.log('Processing', myQueueItem);

    var agencyName = myQueueItem.nc4__agencyName;
    var listName = myQueueItem.nc4__listName;
    var idField = myQueueItem.nc4__idField;

    context.log( 
		'agency=>', 
		agencyName, 
		'list=>', 
		listName, 
		'idField=>', 
		idField,
		'url=>',
		urlForAgency( agencyName ),
		'domain=>',
		domainForAgency( agencyName ),
		'username=>',
		userForAgency( agencyName ),
		'password=>',
		passwordForAgency( agencyName ) );

    var itemJSON = myQueueItem;
    itemJSON.PartitionKey = agencyName + '-' + listName;
    itemJSON.RowKey = myQueueItem[ idField ] + '-' + (new Date()).toISOString();
    context.bindings.tableContent = [ itemJSON ];

    context.done();
};

function settingForAgency(agencyName, settingName)
{
    return process.env[ agencyName + '.' + settingName ];
}

function urlForAgency( agencyName )
{
    return settingForAgency( agencyName, 'url' );
}

function userForAgency( agencyName )
{
    return settingForAgency( agencyName, 'username' )
}

function passwordForAgency( agencyName )
{
    return settingForAgency( agencyName, 'password' );
}

function domainForAgency( agencyName )
{
    return settingForAgency( agencyName, 'domain' );
}

