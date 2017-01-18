var request = require('request'),
  querystring = require('querystring'),
  sharepoint = require('./config'),
  util = require('util'),
  Promise = require('bluebird');

module.exports = {
  getSharePointAppOnlyAccessToken: function (siteUrl, clientId, clientSecret) {
    return new Promise(function (resolve, reject) {
      getRealm(siteUrl).then(function(response){
        //read realm from header
        realm = extractRealmId(response.headers[sharepoint.wwwauthenticate]);
        var body = {};
        body.grant_type = 'client_credentials';
        body.client_id = util.format('%s@%s', clientId, realm);
        body.resource = util.format('%s/%s@%s', sharepoint.sharePointPrinciple, siteUrl, realm).replace('https://', '');
        body.client_secret = clientSecret;
        var formData = querystring.stringify(body);
        var contentLength = formData.length;
        //get app only access token
        request(
          {
            url: util.format(sharepoint.authenticationUrl, realm),
            method: 'POST',
            headers: {
              'Content-Type': 'application/x-www-form-urlencoded',
              'Content-Length': contentLength
            },
            body: formData,
            json: true
          },
          function (error, response, body) {
            if (error || response.statusCode !== 200) {
              return reject(error);
            }
            return resolve(body);
          }
        );
      }).catch(function(error){
        return resolve(error);
      });

    });
  }
}

function getRealm(siteUrl, callback) {
  return new Promise(function (resolve, reject) {
    request({
      url: siteUrl + sharepoint.clientSvcUrl,
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ',
        'Content-Type': 'application/json'
      },
      json: true
    }).on('error', function (error) {
      reject(error);
    }).on('response', function (response) {
      resolve(response);
    });
  });
}

function extractRealmId(header) {
  if (header === null) {
    //throw 'fail to get realm from supplied site Url';
  }
  return header.substring(header.indexOf('realm="') + 7, header.indexOf('",'));
}