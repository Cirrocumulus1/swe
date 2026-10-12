/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";

function requiredProductionValue(name) {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Set ${name} before creating a production manifest.`);
  }
  return value;
}

function requireHttpsUrl(name) {
  const value = requiredProductionValue(name);
  const url = new URL(value);
  if (url.protocol !== "https:") {
    throw new Error(`${name} must use HTTPS.`);
  }
  return url.toString();
}

function createProductionManifest(content) {
  const baseUrl = requireHttpsUrl("ADDIN_BASE_URL").replace(/\/$/, "") + "/";
  const manifest = JSON.parse(content.toString());
  const replaceDevelopmentUrls = (value) => {
    if (Array.isArray(value)) {
      return value.map(replaceDevelopmentUrls);
    }
    if (value && typeof value === "object") {
      return Object.fromEntries(
        Object.entries(value).map(([key, nestedValue]) => [key, replaceDevelopmentUrls(nestedValue)])
      );
    }
    return typeof value === "string" ? value.replaceAll(urlDev, baseUrl) : value;
  };

  const productionManifest = replaceDevelopmentUrls(manifest);
  productionManifest.validDomains = [new URL(baseUrl).host];
  productionManifest.developer = {
    name: requiredProductionValue("DEVELOPER_NAME"),
    websiteUrl: requireHttpsUrl("DEVELOPER_WEBSITE_URL"),
    privacyUrl: requireHttpsUrl("PRIVACY_URL"),
    termsOfUseUrl: requireHttpsUrl("TERMS_OF_USE_URL"),
  };

  return JSON.stringify(productionManifest, null, 2);
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      react: ["react", "react-dom"],
      taskpane: {
        import: ["./src/taskpane/index.tsx", "./src/taskpane/taskpane.html"],
        dependOn: "react",
      },
    },
    output: {
      clean: true,
    },
    resolve: {
      alias: {
        "@fluentui/react-tabster$": require.resolve("@fluentui/react-tabster"),
        "@fluentui/react-positioning$": require.resolve("@fluentui/react-positioning"),
        "@griffel/core$": require.resolve("@griffel/core"),
      },
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: ["ts-loader"],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "react"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.json",
            to: "[name]" + "[ext]",
            transform(content) {
              return dev ? content : createProductionManifest(content);
            },
          },
        ],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      hot: true,
      client: {
        overlay: false,
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      proxy: [
        {
          context: ["/api"],
          target: process.env.API_PROXY_TARGET || "http://localhost:7071",
          changeOrigin: true,
        },
      ],
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
