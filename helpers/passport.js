const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;

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

module.exports = (passport) => {

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
}
