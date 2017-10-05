const $SP = require('sharepointplus');

module.exports = function (context, myQueueItem) {
    var agencyName = myQueueItem.nc4__agencyName;
    var listName = myQueueItem.nc4__listName;
    var idField = myQueueItem.nc4__idField;

    var url = urlForAgency( agencyName );
    var user = userForAgency( agencyName );
    var password = passwordForAgency( agencyName );
    var domain = domainForAgency( agencyName );

    context.log.info( 
		'agency=>', agencyName, 
		'list=>', listName, 
		'idField=>', idField, 
		'url=>', url,
		'domain=>', domain,
		'username=>', user,
		'password=>', password );

    var itemJSON = myQueueItem;
    itemJSON.PartitionKey = agencyName + '-' + listName;
    itemJSON.RowKey = myQueueItem[ idField ] + '-' + (new Date()).toISOString();
    // context.bindings.tableContent = [ itemJSON ];

    var userDefinition = {
        username: user,
        password: password,
        domain: domain
    };

    var sp = $SP().auth( userDefinition );
    var list = sp.list( listName, url );

    processMessage( context, sp, list, idField, itemJSON )

    // context.done();
};

function settingForAgency(agencyName, settingName)
{
    return process.env[ 'agency' + '.' + agencyName + '.' + settingName ];
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

function processMessage( context, _sp, _list, _idField, _msg ) {
        _list
                .get( { fields: '', where: _idField + ' = "' + _msg[ _idField ] + '"' },
                        function( data, error ) {
                                if( error ) { 
                                        context.log.error( 'lookup by ' + idField + ' for value ' + _msg[_idField ] + ' returned error' );
                                        context.log.error( error );
                                        context.done( new Error( 'lookup failed - aborting..' ) );
					return;
                                }       
                                for ( var i = 0; i < data.length; i++ ) {
                                        context.log.info( ' id lookup returned object ' + data[i].getAttribute( _idField ));
                                }       
                                
                                if( data.length === 0 ) {
                                        addToSharePoint( context, _sp, _msg, _idField );
                                } else if ( data.length > 1 ) {
                                        context.done( new Error( 'Only expected one item returned - something is wrong' ) );
                                } else {
                                        updateSharePoint( context, _sp, _msg, _idField );
                                }       
                        } );    
}
                       
function addToSharePoint( context, _sp, _msg, _idField ) {
        context.log( idField + '=' + _msg[_idField] + ' doesnt exists.. adding..' );
        _sp.add( _msg,
                        {
                                error:function(items) {
                                        for (var i=0; i < items.length; i++)
                                                context.log("Add Error '"+items[i].errorMessage+"' with:"+items[i][ _idField ]);
					context.done( new Error( 'Add Error ' + items[0].errorMessage + ' id:' + items[0][ _idField ] ) );
                                },
                                success:function(items) {
                                        for (var i=0; i < items.length; i++)
                                                context.log("Add Success for: (" + _idField  + ":"+items[i][ _idField ] + " )");
					context.done();
                                }
                        }
                );
}

function updateSharePoint( context, _sp, _msg, _idField ) {
        context.log( idField + '=' + _msg[_idField] + ' exists.. updating..' );
        _sp.update( _msg,
                        {
                                where: _idField + ' = "' + _msg[ _idField ] + '"',
                                error:function(items) {
                                        for (var i=0; i < items.length; i++)
                                                context.log("Update Error '"+items[i].errorMessage+"' with:"+items[i][ _idField ]);
					context.done( new Error( 'Update Error ' + items[0].errorMessage + ' id:' + items[0][ _idField ] ) );
                                },
                                success:function(items) {
                                        for (var i=0; i < items.length; i++)
                                                context.log("Update Success for: (" + _idField  + ":"+items[i][ _idField ] + " )");
					context.done();
                                }
                        }
                );
}


