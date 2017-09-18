'use strict';

var MixpanelRetention = require('../lib');
var Fs = require('fs');

var apiKey = process.env.MIXPANEL_API_KEY;
var apiSecret = process.env.MIXPANEL_API_SECRET;

var mixpanelRetention = new MixpanelRetention(apiKey, apiSecret);

mixpanelRetention.retentionReport({
  startDate: '2017-04-30',
  bornEvent: 'Server - Registered User',
  event: 'Session Started',
  unit: 'week'
}).then(function(stream) {
  stream.pipe(Fs.createWriteStream('./output/overall_weekly_retention.xlsx'))
  .on('finish', function () {
      console.log('Successfully generated Overall Weekly Retention spreadsheet.');
  });
});

mixpanelRetention.retentionReport({
  startDate: '2017-04-30',
  bornEvent: 'Server - Registered User',
  event: 'Session Started',
  unit: 'day'
}).then(function(stream) {
  stream.pipe(Fs.createWriteStream('./output/overall_daily_retention.xlsx'))
  .on('finish', function () {
      console.log('Successfully generated Overall Daily Retention spreadsheet.');
  });
});
