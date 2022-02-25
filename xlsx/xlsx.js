const { Readable } = require('stream');
const readXlsxFile = require('read-excel-file/node');

module.exports = function(RED) {

  function bufferToStream(binary) {

    const readableInstanceStream = new Readable({
      read() {
        this.push(binary);
        this.push(null);
      }
    });
    return readableInstanceStream;
  }

  function xlsxNode(config) {
    RED.nodes.createNode(this, config);

    const node = this;
    this.name = config.name;
    this.schema = config.schema;
    this.sheets = config.sheets || 'first';
    this.multi = config.multi || 'one';
    this.parse = config.parse || 'rows';
    this.map = config.map ? JSON.parse(config.map) : null;

    this.on('input', async (msg, send, done) => {
      if (msg.hasOwnProperty('payload')) {
        if (Buffer.isBuffer(msg.payload)) {
          try {

            if (this.sheets === 'first') {
              const options = { sheet: 1 };
              if (this.parse === 'map' && this.map) {
                options.map = this.map;
              }
              const ou = await readXlsxFile(bufferToStream(msg.payload), options);
              msg.payload = ou;
              send(msg);
            } else {
              const sheets = {};
              const sheetNames = await readXlsxFile.readSheetNames(bufferToStream(msg.payload));
  
              for (const sheetName of sheetNames) {
                const options = { sheet: sheetName };
                const ou = await readXlsxFile(bufferToStream(msg.payload), options);
                sheets[sheetName.trim()] = ou;
              }
  
              if (this.multi === 'one') {
                for (const sheet of Object.keys(sheets)) {
                  const m = RED.util.cloneMessage(msg);
                  m.payload = sheets[sheet];
                  send(m);
                }
              } else {
                msg.payload = this.multi === 'array' ? Object.values(sheets) : sheets;
                send(msg);
              }
            }
            done();  
          } catch (error) {
            done(error);
          }
        }
      } else {
        // If no payload just pass it on.
        send(msg);
        done();
      }
    });
  }

  RED.nodes.registerType('xlsx', xlsxNode);
  RED.library.register('maps');
}
