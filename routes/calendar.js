const router = require('express-promise-router')();
const graph = require('../graph.js');
const moment = require('moment-timezone');
const iana = require('windows-iana');
const { body, validationResult } = require('express-validator');
const validator = require('validator');

/* GET /calendar */
router.get('/',
  async function(req, res) {
    console.log (req);

    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      const params = {
        active: { calendar: true }
      };

      // Get the user
      const user = req.app.locals.users[req.session.userId];
      // Convert user's Windows time zone ("Pacific Standard Time")
      // to IANA format ("America/Los_Angeles")
      // Moment needs IANA format
      const timeZoneId = iana.findOneIana(user.timeZone);
      console.log(`Time zone: ${timeZoneId.valueOf()}`);

      // Calculate the start and end of the current week
      // Get midnight on the start of the current week in the user's timezone,
      // but in UTC. For example, for Pacific Standard Time, the time value would be
      // 07:00:00Z
      var startOfWeek = moment.tz(timeZoneId.valueOf()).startOf('week').utc();
      var endOfWeek = moment(startOfWeek).add(7, 'day');
      console.log(`Start: ${startOfWeek.format()}`);

      // Get the access token
      var accessToken;
      try {
        accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
      } catch (err) {
        res.send(JSON.stringify(err, Object.getOwnPropertyNames(err)));
        return;
      }

      if (accessToken && accessToken.length > 0) {
        try {
          // Get the events
          const events = await graph.getCalendarView(
            accessToken,
            startOfWeek.format(),
            endOfWeek.format(),
            user.timeZone);

          res.json(events.value);
        } catch (err) {
          res.send(JSON.stringify(err, Object.getOwnPropertyNames(err)));
        }
      }
      else {
        req.flash('error_msg', 'Could not get an access token');
      }
    }
  }
);

async function getAccessToken(userId, msalClient) {
  // Look up the user's account in the cache
  try {
    const accounts = await msalClient
      .getTokenCache()
      .getAllAccounts();

    const userAccount = accounts.find(a => a.homeAccountId === userId);

    // Get the token silently
    const response = await msalClient.acquireTokenSilent({
      scopes: process.env.OAUTH_SCOPES.split(','),
      redirectUri: process.env.OAUTH_REDIRECT_URI,
      account: userAccount
    });

    return response.accessToken;
  } catch (err) {
    console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
  }
}

module.exports = router;