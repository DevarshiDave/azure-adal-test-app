/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {

  // create
  angular
    .module('meeting-planner-outlook-addin', ['ngRoute', 'AdalAngular'])
    .controller('LoginController', LoginController)
    .controller('HomeController', HomeController)
    .config(['$logProvider', '$routeProvider', 'adalAuthenticationServiceProvider', '$httpProvider', '$locationProvider', function ($logProvider, $routeProvider, adalAuthenticationServiceProvider, $httpProvider, $locationProvider) {
      $locationProvider.html5Mode({
        enabled: true,
        requireBase: false
      });

      // set debug logging to on
      if ($logProvider.debugEnabled) {
        $logProvider.debugEnabled(true);
      }

      $routeProvider
        .when("/", {
          controller: "LoginController",
          templateUrl: "/views/login.html",
          requireADLogin: true
        })
        .when("/home", {
          controller: "HomeController",
          templateUrl: "/views/home.html"
        });


      adalAuthenticationServiceProvider.init({
        // clientId is the identifier assigned to your app by Azure Active Directory.
        clientId: "xxxx-xxxx-xxxx-xxxx",
        endPoints: {
          "https://graph.microsoft.com": "https://graph.microsoft.com"
        },
        extraQueryParameter: 'prompt=admin_consent',
        redirectUri: location.origin,
        cacheLocation: 'localStorage'
      }, $httpProvider);
    }]);

  function LoginController($scope, $http, $q, adalAuthenticationService) {
    console.log('in login', adalAuthenticationService);
    //adalAuthenticationService.acquireTokenRedirect(adalAuthenticationService.config.clientId);
    $http.defaults.useXDomain = true;
    delete $http.defaults.headers.common['X-Requested-With'];

    function getTokenForRequest(resource) {
      var dfd = $q.defer();
      var token = adalAuthenticationService.getCachedToken(resource);
      if (!token) {
        adalAuthenticationService.acquireToken(resource)
          .catch(function (err) {
            console.log(err);
            dfd.reject(error);
          })
          .then(function (data) {
            dfd.resolve(data);
          });
      }
      else {
        dfd.resolve(token);
      }
      return dfd.promise;
    }

    $scope.getGroups = function () {
      getTokenForRequest('https://graph.microsoft.com')
        .catch(function (err) {
          console.log(err);
        })
        .then(function (token) {

          var config = {
            headers: {
              'Authorization': 'Bearer ' + token,
              'Accept': 'application/json'
            }
          };
          $http.get('https://graph.microsoft.com/v1.0/groups?$orderby=displayName', config)
            .catch(function (err) {
              console.log(err);
              adalAuthenticationService.acquireTokenRedirect(adalAuthenticationService.config.endPoints['https://graph.microsoft.com']);
            })
            .then(function (res) {
              $scope.result = res;
            });
        });
    }

    $scope.getSPListItems = function () {
      getTokenForRequest('https://onlinesharepoint2013.sharepoint.com')
        .catch(function (err) {
          console.log(err);
        })
        .then(function (token) {

          var config = {
            headers: {
              'Authorization': 'Bearer ' + token,
              'Accept': 'application/json'
            }
          };
          $http.get("https://onlinesharepoint2013.sharepoint.com/_api/web/lists/getbytitle('Organisations')/items?$select=ID,Title", config)
            .catch(function (err) {
              console.log(err);
            })
            .then(function (res) {
              $scope.result = res;
            });
        });
    }


  }

  /**
   * Home Controller
   */
  function HomeController($scope, $http, adalAuthenticationService) {
    $scope.title = 'Home';
    console.log($scope.title + ' is ready!');

    $scope.run = function () {


      /**
       * Insert your Outlook code here
       */

    }

    var vm = this;
    var adalAuthContext = new AuthenticationContext(adalAuthenticationService.config);

    setTimeout(function () {
      var isCallback = adalAuthContext.isCallback(window.location.hash);
      if (isCallback && !adalAuthContext.getLoginError()) {
        console.log('in handle callback');
        adalAuthContext.handleWindowCallback();
      }
      else {
        var user = adalAuthContext.getCachedUser();
        if (!user) {
          //Log in user
          adalAuthContext.login();
        }
        else {
          console.log(adalAuthContext);
        }
      }
    }, 2500);

    $scope.getGroups = function () {
      // var authHeader = {
      //   Authorization: 'Bearer ' + window.localStorage['mp_token'],
      //   Accept: 'application/json'
      // };
      $http.defaults.useXDomain = true;
      delete $http.defaults.headers.common['X-Requested-With'];

      $http.get('https://graph.microsoft.com/v1.0/groups?$orderby=displayName', {
        headers: { Accept: "application/json;odata.metadata=minimal" }
      }).then(function (res) {
        $scope.result = res;
      });
    }

    // vm.$onInit = function() {
    //   var isCallback = adalAuthContext.isCallback(window.location.hash);
    //   if (isCallback && !adalAuthContext.getLoginError()) {
    //       adalAuthContext.handleWindowCallback();
    //   }
    //   else {
    //       var user = adalAuthContext.getCachedUser();
    //       if (!user) {
    //           //Log in user
    //           adalAuthContext.login();
    //       }
    //   }
    // };
  }

  // when Office has initalized, manually bootstrap the app
  Office.initialize = function () {
    angular.bootstrap(document.body, ['meeting-planner-outlook-addin']);
  };

  setTimeout(function () {
    Office.initialize();
  }, 2500);

})();
