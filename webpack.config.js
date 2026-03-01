const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");

module.exports = {
  entry: {
    taskpane: "./src/taskpane.ts",
    functions: "./src/functions.ts"
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].bundle.js",
    clean: true
  },
  resolve: { extensions: [".ts", ".js"] },
  module: {
    rules: [{ test: /\.ts$/, use: "ts-loader", exclude: /node_modules/ }]
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      chunks: ["taskpane"],
      templateContent: `
<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Managerial Finance Tools</title>
  </head>
  <body>
    <div id="app"></div>
    <script src="taskpane.bundle.js"></script>
  </body>
</html>`
    })
  ],
  devServer: {
    static: { directory: path.join(__dirname, "dist") },
    https: true,
    port: 3000,
    headers: { "Access-Control-Allow-Origin": "*" }
  }
};
