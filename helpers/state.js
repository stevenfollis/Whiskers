require('dotenv').config();
const botbuilderAzure = require('botbuilder-azure');

module.exports = {
  getClient: () => {
    const documentDbOptions = {
      host: process.env.COSMOS_HOST,
      masterKey: process.env.COSMOS_KEY,
      database: process.env.COSMOS_DB,
      collection: process.env.COSMOS_COLLECTION,
    };
    const docDbClient = new botbuilderAzure.DocumentDbClient(documentDbOptions);
    return new botbuilderAzure.AzureBotStorage({ gzipData: false }, docDbClient);
  },
};
