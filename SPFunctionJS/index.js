var graph = require('./graph')();

module.exports = function (context, req) {
  context.log('JavaScript HTTP trigger function processed a request.');

  if (req.query.name || (req.body && req.body.name)) {
    var res;

    graph.getAuth().then(function (token) {
      graph.createItem(token.access_token, req.body.name).then(function () {
        res.body = 'item created';
        context.done(null, res);
      }).catch(function (erro) {
        res.body = 'Error creating item: ' + erro;
        context.done(null, res);
      });
    });
  }
  else {
    res = {
      status: 400,
      body: "Please pass a name on the query string or in the request body"
    };
    context.done(null, res);
  }
};



