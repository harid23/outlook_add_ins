/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return {
    ca: httpsOptions.ca,
    key: httpsOptions.key,
    cert: httpsOptions.cert,
  };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  return {
    devtool: "source-map",

    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      launchevents: "./src/launchevents/launchevents.js",
    },

    output: {
      clean: true,
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
    },

    resolve: {
      extensions: [".js", ".html"],
    },

    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: "babel-loader",
        },
        {
          test: /\.html$/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext]",
          },
        },
      ],
    },

    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext]",
          },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              return dev
                ? content
                : content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),

      new HtmlWebpackPlugin({
        filename: "launchevents.html",
        template: "./src/launchevents/launchevents.html",
        chunks: ["polyfill", "launchevents"],
      }),
    ],

    devServer: {
      static: {
        directory: path.join(__dirname, "dist"),
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: 3000,
    },
  };
};