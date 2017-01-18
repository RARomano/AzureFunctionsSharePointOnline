var graph = require('./graph')();

module.exports = function (context, req) {
  context.log('JavaScript HTTP trigger function processed a request.');

  if (req.query.name || (req.body && req.body.name)) {
    var name = req.query.name || (req.body && req.body.name);

    graph.getAuth().then(function (token) {
      graph.createItem(token.access_token, name).then(function () {
        var res = { body: 'item created' };
        context.done(null, res);
      }).catch(function (erro) {
        var res = {
          body: 'Error creating item: ' + erro,
          status: 500
        };
        context.done(null, res);
      });
    });
  }
  else {
    var res = {
      status: 400,
      body: "Please pass a name on the query string or in the request body"
    };
    context.done(null, res);
  }
};



