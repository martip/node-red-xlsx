# node-red-contrib-xlsx

A [Node-RED](http://nodered.org/) node that converts between a XLSX formatted stream and its JavaScript object representation, in either direction.

This node uses the [read-excel-file](https://www.npmjs.com/package/read-excel-file) and [write-excel-file](https://www.npmjs.com/package/write-excel-file), by catamphetamine.

## Install

Either use the `Node-RED Menu - Manage Palette - Install`, or run the following command in your Node-RED user directory - typically `~/.node-red`

    npm install node-red-contrib-xlsx

## Usage

### XLSX to object

You pass the content of a XLSX file, as a buffer, in the `msg.payload`.

### Object to XLSX

You pass a JSON object, as a string, in the `msg.payload`.