'use strict';

var express = require('express');
var app = express();
var catalyst = require('zcatalyst-sdk-node');
app.use(express.json());
const reader = require('xlsx');//sheet to json
var axios = require('axios');//axios for request

app.get('/parse', (req, res) => {

//initialize catalyst component
var catalystApp = catalyst.initialize(req); 

//read the file
file_read(catalystApp, res);
});

//read the xlsx file
function file_read(catalystApp, res){
	
//Reading our test file
const file = reader.readFile('Data.xlsx')
const sheets = file.SheetNames

//read content stored in array
let data = []

//loop through first two sheets in test file
	for (let i = 0; i < 2; i++) {
		const temp = reader.utils.sheet_to_json(
			file.Sheets[file.SheetNames[i]])
		temp.forEach((res) => {
			//fetch data from rows in excel
			//Todo : Add a counter here to add only 200 rows at a time
			data.push(res)
		})
		//push it into required format
		var data_arr ={"data":data}
		
		connector_catalyst(catalystApp, data_arr, i,res);
		
		//Reset data array to continue filling up for next form
		data=[];
		
	}
}
//catalyst connector : i refers to sheet number
function connector_catalyst(catalystApp, data_arr, i, res){
	var connector = catalystApp.connection({
		ConnectorName: {
	  client_id: {{CLIENT_ID}},
	  client_secret: {{CLIENT_SECRET}},
	  auth_url: 'https://accounts.zoho.com/oauth/v2/token',
	  refresh_url: 'https://accounts.zoho.com/oauth/v2/token',
	  refresh_token: {{REFRESH_TOKEN}}
	 }
	})
	.getConnector('ConnectorName');
	connector.getAccessToken().then((accessToken) => {
		   console.log(accessToken);
			//Invoke add records api with formname (Preferably an array of formnames and invoke name[i])
			if(i==0){
				addrecords(data_arr,"ATL_Crew_List1", accessToken,0, res);
			}else{
				addrecords(data_arr,"ATL_Primary_Cast", accessToken,1, res)
			
					
			}
   }).catch(err => {
	   console.log(err);
	   reject(err);
	  });;		
}


//add records api
function addrecords(data, form_name, accessToken, i,res){
	console.log("Add records API")
		var config = {
	  method: 'post',
	  url: 'https://creator.zoho.in/api/v2/<account_owner_name>/<app_link_name>/form/<form_link_name>+form_name,
	  headers: { 
		'Authorization': 'Zoho-oauthtoken '+accessToken, 
		'Content-Type': 'application/json', 
		  },
	  data : data
	};
	
	axios(config)
	.then(function (response) {
	  console.log(JSON.stringify(response.data));
	  if(i==1){
		  //response message use flags and handle accordingly
		  res.send({
			"message": "Process Completed"
		   });
	  }
	})
	.catch(function (error) {
		//error handling here
	  console.log(error);
	});
}

module.exports = app;
