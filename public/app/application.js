var o365LeaveApp = angular.module("o365LeaveApp", ['ngRoute', 'AdalAngular'])
o365LeaveApp.factory("ShareData", function () {
    return { value: 0 }
});

o365LeaveApp.config(['$routeProvider', '$httpProvider', 'adalAuthenticationServiceProvider', function ($routeProvider, $httpProvider, adalProvider) {
    $routeProvider
           .when('/',
           {
               controller: 'ContactsController',
               templateUrl: '/app/templates/Contacts.html',
               requireADLogin: true
           })
           .when('/contacts/new',
           {
               controller: 'ContactsNewController',
               templateUrl: '/app/templates/ContactsNew.html',
               requireADLogin: true
               
           })
           .when('/contacts/edit',
           {
               controller: 'ContactsEditController',
               templateUrl: '/app/templates/ContactsEdit.html',
               requireADLogin: true

           })
           .when('/contacts/delete',
           {
               controller: 'ContactsDeleteController',
               templateUrl: '/app/templates/Contacts.html',
               requireADLogin: true

           })
           .otherwise({ redirectTo: '/' });
   
    var adalConfig = {
        tenant: '5b532de2-3c90-4e6b-bf85-db0ed9cf5b48',
        clientId: 'e123ba7c-67d7-43c4-8dc2-f9524a14e836',
        extraQueryParameter: 'nux=1',
        endpoints: {
           "https://outlook.office365.com/api/v1.0": "https://outlook.office365.com/"
        }
    };
    adalProvider.init(adalConfig, $httpProvider);
}]);
o365LeaveApp.controller("ContactsController", function ($scope, $q, $location, $http, ShareData, o365CorsFactory) {
    o365CorsFactory.getContacts().then(function (response) {
        $scope.contacts = response.data.value;
    });

    $scope.editUser = function (userName) {
        ShareData.value = userName;
        $location.path('/contacts/edit');
    };
    $scope.deleteUser = function (contactId) {
        ShareData.value = contactId;
        $location.path('/contacts/delete');
    };
});
o365LeaveApp.controller("ContactsNewController", function ($scope, $q, $http, $location, o365CorsFactory) {
    $scope.add = function () {
        var givenName = $scope.givenName
        var surName = $scope.surName
        var email = $scope.email
        contact = {
            "GivenName": givenName, "Surname": surName, "EmailAddresses": [
                    {
                        "Address": email,
                        "Name": givenName
                    }
            ]
        }

        o365CorsFactory.insertContact(contact).then(function () {
            $location.path("/#");
        });
    }
});
o365LeaveApp.controller("ContactsEditController", function ($scope, $q, $http, $location, ShareData, o365CorsFactory) {
    var id = ShareData.value;

    o365CorsFactory.getContact(id).then(function (response) {
        $scope.contact = response.data;
    });
   
    $scope.update = function () {

        var givenName = $scope.contact.GivenName
        var surName = $scope.contact.Surname
        var email = $scope.contact.EmailAddresses[0].Address
        var id = ShareData.value;

        contact = {
            "GivenName": givenName, "Surname": surName, "EmailAddresses": [
                    {
                        "Address": email,
                        "Name": givenName
                    }
            ]
        }
                
        o365CorsFactory.updateContact(contact, id).then(function () {
            $location.path("/#");
        });
    };
});
o365LeaveApp.controller("ContactsDeleteController", function ($scope, $q, $http, $location, ShareData, o365CorsFactory) {
    var id = ShareData.value;
   
    o365CorsFactory.deleteContact(id).then(function () {
        $location.path("/#");
    });
});

o365LeaveApp.factory('o365CorsFactory', ['$http' ,function ($http) {
    var factory = {};
   
    factory.getContacts = function () {
        return $http.get('https://outlook.office365.com/api/v1.0/me/contacts')
    }

    factory.getContact = function (id) {
        return $http.get('https://outlook.office365.com/api/v1.0/me/contacts/'+id)
    }

    factory.insertContact = function (contact) {
       var options = {
            url: 'https://outlook.office365.com/api/v1.0/me/contacts',
            method: 'POST',
            data: JSON.stringify(contact),
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
        };
        return $http(options);
    }

    factory.updateContact = function (contact,id) {
        var options = {
            url: 'https://outlook.office365.com/api/v1.0/me/contacts/'+id,
            method: 'PATCH',
            data: JSON.stringify(contact),
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
        };
        return $http(options);
    }

    factory.deleteContact = function (id) {
        var options = {
            url: 'https://outlook.office365.com/api/v1.0/me/contacts/'+id,
            method: 'DELETE',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
        };
        return $http(options);
    }

    return factory;
}]);

