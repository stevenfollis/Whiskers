const builder = require('botbuilder');
const request = require('request-promise-native');
const emoji = require('node-emoji');
const moment = require('moment');

module.exports = {
  Label: 'Update Status',
  Dialog: [
    (session, args) => {
      console.log('got it');
    }
  ],
};
