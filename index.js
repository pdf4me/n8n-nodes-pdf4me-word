const { Pdf4meWord } = require('./dist/nodes/Pdf4me/Pdf4me.node.js');
const { Pdf4meWordApi } = require('./dist/credentials/Pdf4meApi.credentials.js');

module.exports = {
  nodes: {
    Pdf4meWord,
  },
  credentials: {
    Pdf4meWordApi,
  },
};
