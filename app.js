var auth = require('./auth');
var graph = require('./graph');
var _ = require('underscore');

// Get an access token for the app.
auth.getAccessToken().then(function (token) {

  /// Trocar o usuário
  var userLogin = 'test-createuser2@email.com.br';

  /// Trocar o SKU Name
  var skuName = 'STANDARDWOFFPACK';

  graph.createUser(token, {
    "accountEnabled": true,
    "displayName": "Test-CreateUser",
    "mailNickname": "test-createuser",
    "passwordProfile": {
      "password": "Test1234",
      "forceChangePasswordNextLogin": false
    },
    "userPrincipalName": userLogin
  }).then(function (result) {
    graph.getLicences(token)
        .then(function (licenses) {

          var myLicense = _.where(licenses, {
            skuPartNumber: skuName
          });

          var licenseId = myLicense[0].skuId;

          graph.setUserUsageLocation(token, userLogin, "BR").then(function (result) {

            graph.addLicense(token, userLogin, {
              'addLicenses': [{'skuId': licenseId}],
              'removeLicenses': []
            }).then(function (result) {
              console.log(result);
            });

          });

        }, function (error) {
          console.error('>>> Error getting users: ' + error);
        });
  });

  // graph.getUsers(token)
  //   .then(function (users) {
  //     // Create an event on each user's calendar.
  //     console.log(users);
  //   }, function (error) {
  //     console.error('>>> Error getting users: ' + error);
  //   });
}, function (error) {
  console.error('>>> Error getting access token: ' + error);
});