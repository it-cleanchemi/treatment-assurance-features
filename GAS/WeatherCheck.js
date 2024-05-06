function weatherTrigger() {
  // First, delete existing triggers for the 'delayedFunction'
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'postWeatherUpdate') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }

  // Now set up a new trigger
  ScriptApp.newTrigger('postWeatherUpdate')
           .timeBased()
           .after(60 * 60 * 1000) // Delay for 10 minutes
           .create();
}

function postWeatherUpdate() {
  
  var name = SS.getName();
  var lat = SS.getSheetByName('Reference').getRange("H2").getValue();
  var long = SS.getSheetByName('Reference').getRange("H3").getValue();
  
  
  if (lat===""||long===""){
    return
  }
  
  var formattedLat = parseFloat(lat.toFixed(4));  // Converts the string back to a float
  var formattedLong = parseFloat(long.toFixed(4)); // Converts the string back to a float

  var city = name.substring(0, 5);
  var url = 'https://api.weather.gov/points/'+formattedLat+','+formattedLong ;



 var options = {
    'method': 'get',
    'headers': {
      'User-Agent': 'my-google-apps-script (v.martysevich@cleanchemi.com)' // Change this to your email
    }
  };

 // Fetch the office and grid coordinates from the NWS API
  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  // Extracting Zone ID for alerts
  var zoneId = json.properties.forecastZone.split('/').pop();
  var alerts = fetchAlerts(zoneId, options);

  // Fetch current weather and forecast URLs from properties
  var forecastUrl = json.properties.forecast; // URL for the hourly forecast
  var hourlyUrl = json.properties.forecastHourly;

  // Fetch the 12-hour forecast
  var forecastResponse = UrlFetchApp.fetch(forecastUrl, options);
  var hourlyResopince = UrlFetchApp.fetch(hourlyUrl, options);
  var forecastJson = JSON.parse(forecastResponse.getContentText());
  var hourlyJson = JSON.parse(hourlyResopince.getContentText());

  // Using the first period for current detailed forecast
  var currentWeather = forecastJson.properties.periods[0].detailedForecast;

  var periods = hourlyJson.properties.periods.slice(0, 12); // Get the first 12 periods for 12 hours

  // Creating a card payload for posting to Google Chat
  var cardPayload = createWeatherCard(city, currentWeather, alerts, periods);
  var payload = JSON.stringify(cardPayload);

  

  var webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAAApEyy8XY/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=cc0XVeeTaP9QLBbrDecNG76cqBxHppp_SAROy6MTAvg';
   UrlFetchApp.fetch(webhookUrl, {
    'method': 'post',
    'contentType': 'application/json',
    'payload': payload
  });
}

// Prepare card

function createWeatherCard(city, currentWeather, alerts, periods) {
    var sections = [];

    // Current weather section
    sections.push({
        "widgets": [{
            "textParagraph": {
                "text": "<b>Current weather:</b><br>" + currentWeather
            }
        }]
    });

    // Alerts section
    if (alerts.length > 0) {
        var alertText = alerts.map(alert => "<b>" + alert.properties.event + ":</b><br>" + alert.properties.headline).join('<br>');
        sections.push({
            "widgets": [{
                "textParagraph": {
                    "text": "<b>Active Alerts:</b><br>" + alertText
                }
            }]
        });
    } else {
        sections.push({
            "widgets": [{
                "textParagraph": {
                    "text": "<b>Active Alerts:</b><br>No active alerts"
                }
            }]
        });
    }

    // 12-hour forecast section
    var forecastText = periods.map(period =>
        `${period.startTime.substring(11, 16)} - ${period.shortForecast}, ${period.temperature}°${period.temperatureUnit}`
    ).join('<br>');

    sections.push({
        "widgets": [{
            "textParagraph": {
                "text": "<b>12-Hour Forecast:</b><br>" + forecastText
            }
        }]
    });

    // Create the card
    var card = {
        "cards": [{
            "header": {
                "title": "Weather Update for " + city,
                "subtitle": "Provided by: National Weather Service",
                "imageUrl": "https://fonts.gstatic.com/s/e/notoemoji/15.1/1f326_fe0f/512.png=s40",
                "imageStyle": "IMAGE"
            },
            "sections": sections
        }]
    };

    return card;
}

function fetchAlerts(zoneId, options) {
  var alertsUrl = "https://api.weather.gov/alerts/active/zone/" + zoneId; // Correct endpoint for active alerts by zone

  try {
    var alertsResponse = UrlFetchApp.fetch(alertsUrl, {
      ...options,
      muteHttpExceptions: true  // This prevents the script from throwing an exception
    });
    if (alertsResponse.getResponseCode() === 200) {
      var alertsJson = JSON.parse(alertsResponse.getContentText());
      return alertsJson.features;
    } else {
      console.log('No active alerts or unable to fetch alerts:', alertsResponse.getContentText());
      return [];  // Return an empty array if the API call did not succeed
    }
  } catch (error) {
    console.error('Error fetching alerts:', error);
    return [];  // Return an empty array in case of errors
  }
}
