'use strict';

var MixpanelRetention = require('./MixpanelRetention');

var initializer = function(apiKey, apiSecret) {
  return new MixpanelRetention(apiKey, apiSecret);
};

initializer.MixpanelRetention = MixpanelRetention;

module.exports = initializer;
