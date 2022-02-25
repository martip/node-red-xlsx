# node-red-contrib-xlsx

A [Node-RED](http://nodered.org/) node that parses a XLSX file and converts it to a JavaScript representation (object or array).

This node uses the [read-excel-file](https://www.npmjs.com/package/read-excel-file) package, by catamphetamine.

## Install

Either use the `Node-RED Menu - Manage Palette - Install`, or run the following command in your Node-RED user directory - typically `~/.node-red`

    npm install node-red-contrib-xlsx

## Usage

You pass the content of a XLSX file, as a buffer, in the `msg.payload`.

Depending on how it's configured, the node outputs a single message or a message for each parsed sheet, containing arrays or objects.