const $SP = require('sharepointplus');

module.exports = function (context, myQueueItem) {
        var agencyName = myQueueItem.nc4__agencyName;
        var listName = myQueueItem.nc4__listName;
        var idField = myQueueItem.nc4__idField;

        var url = urlForAgency(agencyName);
        var user = userForAgency(agencyName);
        var password = passwordForAgency(agencyName);
        var domain = domainForAgency(agencyName);

        context.log.info(
                'agency=>', agencyName, 'list=>', listName, 'idField=>', idField,
                'url=>', url, 'domain=>', domain, 'username=>', user, 'password=>', password);

        var outputBinding = cloneForOutputBinding( context, agencyName, listName, idField, myQueueItem);
        context.bindings.tableContent = [outputBinding];
                
        var sharepointObj = JSON.parse(JSON.stringify(myQueueItem));
        delete sharepointObj.nc4__agencyName;
        delete sharepointObj.nc4__listName;
        delete sharepointObj.nc4__idField;
        context.log('sharepointObj', sharepointObj);

        var userDefinition = {
                username: user,
                password: password,
                domain: domain
        };

        try {
                var sp = $SP().auth(userDefinition);
                var list = sp.list(listName, url);
                processMessage(context, sp, list, idField, sharepointObj);
        } catch (ex) {
                context.log.error('exception handler triggered', ex);
                outputBinding.nc4_error = ex;
                context.done();
        }

        // context.done();
};

function cloneForOutputBinding( context, agencyName, listName, idField, msg ) {
        var outputBinding = JSON.parse(JSON.stringify(myQueueItem));
        outputBinding.PartitionKey = agencyName + '-' + listName;
        outputBinding.RowKey = myQueueItem[idField] + '-' + (new Date()).toISOString();
        context.log('outputBinding', outputBinding);
        return outputBinding;
}

function processMessage(context, _sp, _list, _idField, _msg) {
        context.log('processMessage, entry()');
        _list
                .get({
                        fields: '',
                        where: _idField + ' = "' + _msg[_idField] + '"'
                },
                function (data, error) {
                        context.log('get callback triggered');
                        if (error) {
                                context.log.error('lookup by ' + idField + ' for value ' + _msg[_idField] + ' returned error');
                                context.binding.tableContent.nc4__error = error;
                                context.done();
                                return;
                        } else {

                                for (var i = 0; i < data.length; i++) {
                                        context.log.info(' id lookup returned object ' + data[i].getAttribute(_idField));
                                }

                                if (data.length === 0) {
                                        context.log.error('data.length was 0');
                                        addToSharePoint(context, _sp, _msg, _idField);
                                } else if (data.length > 1) {
                                        context.log.error('data.length was > 1');
                                        context.binding.tableContent.nc4__error = 'something wrong';
                                        context.done('Only expected one item returned - something is wrong', _msg);
                                } else {
                                        context.log('data.length was 1');
                                        updateSharePoint(context, _sp, _msg, _idField);
                                }
                        }
                });
}

function addToSharePoint(context, _sp, _msg, _idField) {
        context.log(_idField + '=' + _msg[_idField] + ' doesnt exists.. adding..');
        _sp.add(_msg, {
                error: function (items) {
                        context.log.error('addToSharePoint:error() triggered');
                        for (var i = 0; i < items.length; i++)
                                context.log.error("Add Error '" + items[i].errorMessage + "' with:" + items[i][_idField]);
                        context.binding.tableContent.nc4__error = items[0].errorMessage;
                        context.done();
                },
                success: function (items) {
                        context.log('addToSharePoint:success() triggered');
                        for (var i = 0; i < items.length; i++)
                                context.log("Add Success for: (" + _idField + ":" + items[i][_idField] + " )");
                        context.done();
                }
        });
}


function updateSharePoint(context, _sp, _msg, _idField) {
        context.log(_idField + '=' + _msg[_idField] + ' exists.. updating..');
        context.log(_msg);
        _sp.update(_msg, {
                where: _idField + ' = "' + _msg[_idField] + '"',
                error: function (items) {
                        context.log.error('updateToSharePoint:error() triggered');
                        for (var i = 0; i < items.length; i++)
                                context.log.error("Update Error '" + items[i].errorMessage + "' with:" + items[i][_idField]);
                        context.binding.tableContent.nc4__error = items[0].errorMessage;
                        context.done();
                },
                success: function (items) {
                        context.log('updateToSharePoint:success() triggered');
                        for (var i = 0; i < items.length; i++)
                                context.log("Update Success for: (" + _idField + ":" + items[i][_idField] + " )");
                        context.done();
                }
        });
}

function settingForAgency(agencyName, settingName) {
        return process.env['agency' + '.' + agencyName + '.' + settingName];
}

function urlForAgency(agencyName) {
        return settingForAgency(agencyName, 'url');
}

function userForAgency(agencyName) {
        return settingForAgency(agencyName, 'username')
}

function passwordForAgency(agencyName) {
        return settingForAgency(agencyName, 'password');
}

function domainForAgency(agencyName) {
        return settingForAgency(agencyName, 'domain');
}

