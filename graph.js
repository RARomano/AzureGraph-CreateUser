var request = require('request');
var Q = require('q');
var config = require('./config');

// The graph module object.
var graph = {};

// @name getUsers
// @desc Makes a request to the Microsoft Graph for all users in the tenant.
graph.getUsers = function (token) {
  var deferred = Q.defer();

  // Make a request to get all users in the tenant. Use $select to only get
  // necessary values to make the app more performant.
  request.get('https://graph.windows.net/'+config.tenantId+'/users?$select=id,displayName', {
    auth: {
      bearer: token
    }
  }, function (err, response, body) {
    var parsedBody = JSON.parse(body);

    if (err) {
      deferred.reject(err);
    } else if (parsedBody.error) {
      deferred.reject(parsedBody.error.message);
    } else {
      // The value of the body will be an array of all users.
      deferred.resolve(parsedBody.value);
    }
  });

  return deferred.promise;
};

graph.getLicences = function (token) {
  var deferred = Q.defer();

  // Make a request to get all users in the tenant. Use $select to only get
  // necessary values to make the app more performant.
  request.get('https://graph.windows.net/'+config.tenantId+'/subscribedSkus?api-version=1.6', {
    auth: {
      bearer: token
    }
  }, function (err, response, body) {
    var parsedBody = JSON.parse(body);

    if (err) {
      deferred.reject(err);
    } else if (parsedBody.error) {
      deferred.reject(parsedBody.error.message);
    } else {
      // The value of the body will be an array of all users.
      deferred.resolve(parsedBody.value);
    }
  });

  return deferred.promise;
};

graph.addLicense = function (token, user, userData) {
  var deferred = Q.defer();

  var postData = {
    auth: { bearer: token },
    json: userData,
    url: 'https://graph.windows.net/'+config.tenantId+'/users/'+user+'/assignLicense?api-version=1.6'
  };

  // Make a request to get all users in the tenant. Use $select to only get
  // necessary values to make the app more performant.
  request.post(postData, function (err, response, body) {
    if (err) {
      deferred.reject(err);
    } else {
      // The value of the body will be an array of all users.
      deferred.resolve(body);
    }
  });

  return deferred.promise;
};

graph.setUserUsageLocation = function (token, user, usageLocation) {
  var deferred = Q.defer();

  var postData = {
    auth: { bearer: token },
    json: {
      'usageLocation': usageLocation
    },
    url: 'https://graph.windows.net/'+config.tenantId+'/users/'+user+'?api-version=1.6'
  };

  // Make a request to get all users in the tenant. Use $select to only get
  // necessary values to make the app more performant.
  request.patch(postData, function (err, response, body) {
    if (err) {
      deferred.reject(err);
    } else {
      // The value of the body will be an array of all users.
      deferred.resolve(body);
    }
  });

  return deferred.promise;
};

graph.createUser = function (token, userData) {
  var deferred = Q.defer();

  var postData = {
    auth: { bearer: token },
    json: userData,
    url: 'https://graph.windows.net/'+config.tenantId+'/users?api-version=1.6'
  };

  // Make a request to get all users in the tenant. Use $select to only get
  // necessary values to make the app more performant.
  request.post(postData, function (err, response, body) {
    if (err) {
      deferred.reject(err);
    } else {
      // The value of the body will be an array of all users.
      deferred.resolve(body);
    }
  });

  return deferred.promise;
};

module.exports = graph;