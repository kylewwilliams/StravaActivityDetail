// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Strava')
    .addItem('Get Strava activities', 'getStravaActivityDetails')
    .addToUi();
}

// Get athlete activity data and activity details
function getStravaActivityDetails() {
  
  // get the sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');

  // get the latest timestamp
  var lastrow = sheet.getLastRow();
  var lastdate = SpreadsheetApp.getActiveSheet().getRange(lastrow, 2).getValue();
  var lastactid = SpreadsheetApp.getActiveSheet().getRange(lastrow, 10).getValue();
  Logger.log(lastdate);
  Logger.log(lastactid);
  var lastdateepoch = (((new Date(lastdate)).getTime()) /1000);
  //Logger.log(lastdateepoch);


  // call the Strava API to retrieve basic activity data
  var actdata = callStravaAPIact(lastdateepoch);
  var actdatalength = actdata.length;

  // loop over activity data and add the activity ID number to actids array 
  var actids = [];
  for (var j = 0; j < actdatalength; j++) {
      actids.push( 
        ''+actdata[j].id+'',
      );
    } // end for
  //Logger.log(actids);

  var actidslength = actids.length;
  //Logger.log(actidslength);

  // loop through the activity IDs and feed them to the callStravaAPI, which retrieves activity details
  var actidstart = [];
  if (actids[0] == lastactid){
     actidstart = 1;
  } else {
    actidstart = 0;
  }
  //Logger.log (actidstart);
  for (var i = actidstart; i < actidslength; i++) {
    var actid = actids[i];
    //Logger.log(actid);
  
  // call the Strava API to retrieve activity details data
  var data = callStravaAPI(actid);

  //Logger.log(data.best_efforts);
  var besteffortsdata = [];
  if (data.best_efforts == null){
    besteffortsdata = [];
  } else {
    // figure out how many best efforts data points there are
    var becount = data.best_efforts.length;

    // get each best efforts data
    for (var counter = 0; counter < becount; counter = counter + 1){
      besteffortsdata.push( 
        data.best_efforts[counter].name,
        data.best_efforts[counter].elapsed_time,
        data.best_efforts[counter].distance,
        data.best_efforts[counter].start_index,
        data.best_efforts[counter].end_index,
        data.best_efforts[counter].pr_rank)
        ;
    } // end for
   } // end else
   //Logger.log(besteffortsdata);
   //Logger.log(data.type);
  
  // empty array to hold activity data
  var stravaData = [];
  if(data.type == 'Run' || data.type == 'Ride'){
      stravaData = stravaData.concat([
        data.type,
        data.start_date_local,
        data.name,
        data.workout_type,
        data.distance,
        data.calories,
        data.total_elevation_gain,
        data.moving_time,
        data.elapsed_time,
        data.id,
        data.athlete.id,
        data.timezone,
        data.location_country,
        data.start_latitude,
        data.start_longitude,
        data.achievement_count,
        data.kudos_count,
        data.comment_count,
        data.athlete_count,
        data.photo_count,
        data.map.id,
        data.map.polyline,
        data.map.summary_polyline,
        data.commute,
        data.manual,
        data.gear.id,
        data.gear.name,
        data.device_name,
        data.average_speed,
        data.average_heartrate,
        data.max_heartrate,
        data.elev_high,
        data.elev_low,     
      ]);
    } //end if
    else{
      stravaData = stravaData.concat([
        data.type,
        data.start_date_local,
        data.name,
        data.workout_type,
        data.distance,
        data.calories,
        data.total_elevation_gain,
        data.moving_time,
        data.elapsed_time,
        data.id,
        data.athlete.id,
        data.timezone,
        data.location_country,
        data.start_latitude,
        data.start_longitude,
        data.achievement_count,
        data.kudos_count,
        data.comment_count,
        data.athlete_count,
        data.photo_count,
        data.map.id,
        data.map.polyline,
        data.map.summary_polyline,
        data.commute,
        data.manual,
        'null',
        'null',
        data.device_name,
        data.average_speed,
        data.average_heartrate,
        data.max_heartrate,
        data.elev_high,
        data.elev_low,     
      ]);

    } //end else

  // add best efforts data to the activity data
    stravaData = stravaData.concat(besteffortsdata);

  // combine into array with second square brackets
    var stravaData1 = [];
    stravaData1.push(stravaData);
  
  // paste the values into the Sheet
  sheet.getRange(sheet.getLastRow() + 1, 1, stravaData1.length, stravaData1[0].length).setValues(stravaData1);
}
}//endfor

// call the Strava API for activity details
function callStravaAPI(actid) {
  
  // set up the service
  var service = getStravaService();

  if (service.hasAccess()) {
    //Logger.log('Activity detail app has access.');
    
    var endpoint = 'https://www.strava.com/api/v3/activities/'+actid;
    var params = '?include_all_efforts=true';

    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };
    
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true
    };
    
    var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
    //Logger.log(response);
    
    return response;  
 
  }
  else {
    Logger.log("Activity detail app has no access yet.");
    
    // open this url to gain authorization from github
    var authorizationUrl = service.getAuthorizationUrl();
    
    Logger.log("Open the following URL and re-run the script: %s",
        authorizationUrl);
  }
}  

// call the Strava API for activity IDs
function callStravaAPIact(lastdateepoch) {
  
  // set up the service
  var service = getStravaService();
  
  if (service.hasAccess()) {
    Logger.log('Activity data app has access.');
    
    var endpoint = 'https://www.strava.com/api/v3/athlete/activities';
    var params = '?after='+lastdateepoch+'&per_page=99';

    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };
    
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true
    };
    
    var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options).getContentText())
    //Logger.log(response);
  
    
    return response;  
 
  }
  else {
    Logger.log("Activity data app has no access yet.");
    
    // open this url to gain authorization from github
    var authorizationUrl = service.getAuthorizationUrl();
    
    Logger.log("Open the following URL and re-run the script: %s",
        authorizationUrl);
  }
}

//Use
//{{athlete?.id}}
//to make the code null-safe for async loaded data.
