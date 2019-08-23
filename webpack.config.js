const path = require("path");
const HtmlPlugin = require("html-webpack-plugin");


module.exports = {
  entry: path.resolve(__dirname, "src/index.js"),

  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js"
  },

  plugins: [
    new HtmlPlugin({
      template: path.resolve(__dirname, "src/index.html")
    })
  ],

  node: {
    fs: "empty"
  }
}