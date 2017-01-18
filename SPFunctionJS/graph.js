var request = require('request'),
  config = require('./config'),
  Promise = require('bluebird'),
  sharepoint = require('./sharepoint-apponly');

module.exports = function () {

  return {
    getAuth: function () {
      return sharepoint.getSharePointAppOnlyAccessToken(config.siteUrl,
        config.clientId,
        config.clientSecret);
    },

    createItem: function (token, title) {
      var url = config.siteUrl + "/_api/Web/Lists/GetByTitle('"+config.listUrl+"')/Items";
      return new Promise(function (resolve, reject) {
        var postData = {
          auth: {bearer: token},
          json: {
            'Title': title
          },
          url: url,
          headers: {
            accept: 'application/json'
          }
        };

        // Make a request to get all users in the tenant. Use $select to only get
        // necessary values to make the app more performant.
        request.post(postData, function (err, response, body) {
          if (err || (body && body["odata.error"])) {
            reject(err || body);
          } else {
            // The value of the body will be an array of all users.
            resolve(body);
          }
        });

      });
    }
  }
};