const passport = require('passport-restify');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const crypto = require('crypto');
const builder = require('botbuilder');

const strategy = {
  redirectUrl: `${process.env.AUTHBOT_CALLBACKHOST}/api/OAuthCallback`,
  realm: process.env.MICROSOFT_REALM,
  clientID: process.env.MICROSOFT_CLIENT_ID,
  clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
  validateIssuer: false,
  allowHttpForRedirectUrl: true,
  issuer: null,
  identityMetadata: `https://login.microsoftonline.com/${process.env.MICROSOFT_REALM}/.well-known/openid-configuration`,
  skipUserProfile: true,
  scope: null,
  loggingLevel: 'error',
  nonceLifetime: null,
  responseType: 'code id_token',
  responseMode: 'query',
  passReqToCallback: true,
};

module.exports = (server, bot) => {
  server.get('/login', (req, res, next) => {
    passport.authenticate('azuread-openidconnect', { failureRedirect: '/login', customState: req.query.address, resourceURL: process.env.MICROSOFT_RESOURCE_GRAPH },
      (err, user) => {
        if (err) {
          console.log(err);
          return next(err);
        }
        if (!user) {
          return res.redirect('/login');
        }
        req.logIn(user, (logInError) => {
          if (logInError) {
            return next(logInError);
          }
          return res.send(`Howdy ${req.user.displayName}`);
        });
      })(req, res, next);
  });

  server.get('/api/OAuthCallback/',
    passport.authenticate('azuread-openidconnect', { failureRedirect: '/login', resourceURL: process.env.MICROSOFT_RESOURCE_GRAPH }),
    (req, res) => {
      const address = JSON.parse(req.query.state);
      const magicCode = crypto.randomBytes(4).toString('hex');
      const messageData = {
        magicCode,
        accessToken: req.authInfo.accessToken,
        refreshToken: req.authInfo.refreshToken,
        name: req.user.displayName,
        upn: req.user.upn,
      };

      const continueMsg = new builder.Message().address(address).text(JSON.stringify(messageData));

      bot.receive(continueMsg.toMessage());

      // Send a page in the browser displaying the magic code
      const page = `<!doctype html><style>body{text-align: center; padding: 150px;}h1{font-size: 40px;}body{font: 20px Helvetica, sans-serif; color: #333;}article{display: block; text-align: left; width: 650px; margin: 0 auto;}code{font-family:Consolas,monaco,monospace;}</style><article> <h1>Welcome ${req.user.displayName}!</h1> <div> <p>Please copy this code and paste it back to your chat so your authentication can complete:</p><code>${magicCode}</code></div></article>`;
      res.end(page);
    });

  passport.serializeUser((user, done) => {
    done(null, user);
  });

  passport.deserializeUser((id, done) => {
    done(null, id);
  });

  passport.use(new OIDCStrategy(strategy,
    (req, iss, sub, profile, accessToken, refreshToken, done) => {
      if (!profile.oid) {
        return done(new Error('No oid found'), null);
      }
      // asynchronous verification, for effect...
      process.nextTick(() => {
        const tokens = { accessToken, refreshToken };
        return done(null, profile, tokens);
      });
    }));
};
