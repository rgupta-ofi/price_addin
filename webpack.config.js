/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");
const Dotenv = require("dotenv-webpack");
const devCerts = require("office-addin-dev-certs");

const urlDev = "https://localhost:3000/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const isDev = options.mode === "development";

  const config = {
    devtool: isDev ? "inline-source-map" : "source-map",
    entry: {
      functions: "./src/functions/functions.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".js", ".json"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: [isDev ? "style-loader" : MiniCssExtractPlugin.loader, "css-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif|svg|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext]",
          },
        },
      ],
    },
    plugins: [
      new Dotenv({ systemvars: true }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["functions"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "src/assets", to: "assets", noErrorOnMissing: true },
          { from: "manifest.xml", to: "manifest.xml" },
          { from: "src/functions/functions.json", to: "functions.json" },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, PATCH, OPTIONS",
        "Access-Control-Allow-Headers": "X-Requested-With, content-type, Authorization"
      },
      server: {
        type: "https",
        options: isDev ? await getHttpsOptions() : {},
      },
      port: 3000,
      allowedHosts: "all",
      static: [
        {
          directory: path.join(__dirname, "dist"),
        },
        {
          directory: path.join(__dirname, "src", "assets"),
          publicPath: "/assets",
        }
      ],
    },
  };

  if (!isDev) {
    config.plugins.push(new MiniCssExtractPlugin({ filename: "[name].css" }));
  }

  return config;
};
