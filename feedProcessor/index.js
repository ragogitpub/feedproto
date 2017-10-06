const $SP = require('sharepointplus');

module.exports = function (context, myQueueItem) {
        context.log(`Dequeue count: ${context.bindingData.dequeueCount}`, myQueueItem);
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

        var outputBinding = cloneForOutputBinding(context, agencyName, listName, idField, myQueueItem);
        context.bindings.tableContent = [outputBinding];

        var sharepointObj = cloneForSharePoint(context, myQueueItem);

        var userDefinition = {
                username: user,
                password: password,
                domain: domain
        };

        try {
                var sp = $SP().auth(userDefinition);
                var list = sp.list(listName, url);
                processMessage(context, sp, list, idField, sharepointObj);
                //context.done();
        } catch (ex) {
                context.log.error('exception handler triggered', ex);
                context.done('exception handler triggered');
        }

        // context.done();
};

function cloneForOutputBinding(context, agencyName, listName, idField, msg) {
        var outputBinding = JSON.parse(JSON.stringify(msg));
        outputBinding.PartitionKey = agencyName + '-' + listName;
        outputBinding.RowKey = msg[idField] + '-' + (new Date()).toISOString();
        //context.log('outputBinding', outputBinding);
        return outputBinding;
}

function cloneForSharePoint(context, msg) {
        var sharepointObj = JSON.parse(JSON.stringify(msg));
        delete sharepointObj.nc4__agencyName;
        delete sharepointObj.nc4__listName;
        delete sharepointObj.nc4__idField;
        //context.log('sharepointObj', sharepointObj);
        return sharepointObj;
}

function processMessage(context, _sp, _list, _idField, _msg) {
        _list
                .get({
                                fields: '',
                                where: _idField + ' = "' + _msg[_idField] + '"'
                        },
                        function (data, error) {
                                if (error) {
                                        context.log.error('lookup by ' + idField + ' for value ' + _msg[_idField] + ' returned error');
                                        context.done(error);
                                        return;
                                } else {
                                        if (data.length === 0) {
                                                addToSharePoint(context, _sp, _msg, _idField);
                                        } else if (data.length > 1) {
                                                context.log.error('data.length was > 1');
                                                context.done('Only expected one item returned - something is wrong');
                                        } else {
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
                        try {
                                context.log.error(items[0].errorMessage);
                                context.done(items[0].errorMessage);
                        } catch(ex) {
                                context.log.error(ex);
                                context.done();
                        }
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
        _sp.update(_msg, {
                where: _idField + ' = "' + _msg[_idField] + '"',
                error: function (items) {
                        context.log.error('updateToSharePoint:error() triggered');
                        handleError( context, items[0].errorMessage);
                },
                success: function (items) {
                        context.log('updateToSharePoint:success() triggered');
                        for (var i = 0; i < items.length; i++)
                                context.log("Update Success for: (" + _idField + ":" + items[i][_idField] + " )");
                        context.done();
                }
        });
}

function handleError(context,errorMessage) {
        try {
                context.log.error(errorMessage);
                var originalBinding = JSON.parse(JSON.stringify(context.bindings.tableContent[0]));
                var errorBinding = JSON.parse(JSON.stringify(context.bindings.tableContent[0]));
                originalBinding.nc4__error = errorMessage;
                errorBinding.nc4__error = errorMessage;
                errorBinding.PartitionKey = errorBinding.PartitionKey + '-Errors';
                context.bindings.tableContent[0] = originalBinding;
                context.bindings.tableContent[1] = errorBinding;
                context.bindings.emailErrorMessage = [{
                        "personalizations": [ { "to": [ { "email": "rajesh.goswami@nc4.com" } ] } ],
                       "content": [{
                           "type": 'text/plain',
                           "value": JSON.stringify(errorBinding)
                       }]
                   }];
                context.done();
        } catch(ex) {
                context.log.error(ex);
                context.done();
        }
}

function settingForAgency(agencyName, settingName) {
        return process.env['agency' + '.' + agencyName + '.' + settingName];
}

function urlForAgency(agencyName) {
        return settingForAgency(agencyName, 'url');
}

function userForAgency(agencyName) {
        return settingForAgency(agencyName, 'username');
}

function passwordForAgency(agencyName) {
        return settingForAgency(agencyName, 'password');
}

function domainForAgency(agencyName) {
        return settingForAgency(agencyName, 'domain');
}