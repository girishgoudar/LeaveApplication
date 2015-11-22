var express = require('express');
var router = express.Router();
var authContext = require('adal-node').AuthenticationContext;
var authHelper = require('../authHelper.js');
var https = require('https')

/* GET home page. */
router.get('/', function(req, res, next) {
  if(req.cookies.TOKEN_CACHE_KEY === undefined){
    res.redirect('/login')
  }
  else{
    authHelper.getTokenFromRefreshToken('https://api.office.com/discovery/',req.cookies.TOKEN_CACHE_KEY,function(token){
      var endpoint = "https://api.office.com/discovery/v1.0/me/services('RootSite@O365_SHAREPOINT')?$select=serviceResourceId";
      getJson('api.office.com',endpoint,token.accessToken, function(json){
          res.cookie(authHelper.TENANT_URL, JSON.parse(json).serviceResourceId) 
          res.render('index', { title: 'Discovery services' });
       });
     });
  }
 });

router.get('/login', function(req, res, next) {
  if(req.query.code === undefined){
    res.render('login',{title :'login to office 365', authRedirect:authHelper.getAuthUrl('https://api.office.com/discovery/')})
  }
  else{
    authHelper.getTokenFromCode('https://api.office.com/discovery/',req.query.code,function(token){
     res.cookie(authHelper.TOKEN_CACHE_KEY , token.refreshToken);
     res.cookie(authHelper.TENANT_CAHCE_KEY, token.tenantId) 
     res.redirect('/')
    });
  }
});

function getJson(host,endpoint, token, callback) {
  var options = {
    host: host, 
    path: endpoint, 
    method: 'GET', 
    headers: { 
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + token 
      }
    };

  https.get(options, function(response) {
    var body = "";
    response.on('data', function(d) {
      body += d;
    });
    response.on('end', function() {
      callback(body);
    });
    response.on('error', function(e) {
      callback(null);
    });
  });
  
  
};



module.exports = router;



