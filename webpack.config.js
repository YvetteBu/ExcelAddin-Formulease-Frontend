/* eslint-disable no-undef */
import webpack from "webpack";
import Dotenv from "dotenv-webpack";
import CopyWebpackPlugin from "copy-webpack-plugin";
import HtmlWebpackPlugin from "html-webpack-plugin";
import path from "path";
import MiniCssExtractPlugin from "mini-css-extract-plugin";

const urlDev = "https://localhost:3000/";
const urlProd = "https://excel-addin-formulease-rjzjyrndk-yvettebus-projects.vercel.app/"; // Production deployment URL

const config = {
    devtool: "source-map",
    entry: {
        polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
        taskpane: "./src/taskpane/taskpane.js"
    },
    output: {
        clean: true,
        publicPath: "/",
    },
    resolve: {
        extensions: [".html", ".js"]
    },
    module: {
        rules: [
            {
                test: /\.js$/,
                exclude: /node_modules/,
                use: {
                    loader: "babel-loader",
                    options: {
                        presets: ["@babel/preset-env"]
                    }
                }
            },
            {
                test: /\.html$/,
                exclude: /node_modules/,
                use: "html-loader"
            },
            {
                test: /\.(png|jpg|jpeg|gif|ico)$/,
                type: "asset/resource",
                generator: {
                    filename: "assets/[name][ext][query]"
                }
            },
            {
                test: /\.css$/i,
                use: [MiniCssExtractPlugin.loader, "css-loader"],
            }
        ]
    },
    plugins: [
        new CopyWebpackPlugin({
            patterns: [
                {
                    from: "assets/*",
                    to: "assets/[name][ext][query]"
                },
                {
                    from: "manifest*.xml",
                    to: "[name][ext]",
                    transform(content) {
                        return content;
                    }
                }
            ]
        }),
        new HtmlWebpackPlugin({
            filename: "taskpane.html",
            template: "./src/taskpane/taskpane.html",
            chunks: ["polyfill", "taskpane"]
        }),
        new webpack.ProvidePlugin({
            Promise: ["es6-promise", "Promise"]
        }),
        new Dotenv({
            systemvars: true
        }),
        new MiniCssExtractPlugin({
            filename: "[name].css"
        }),
    ],
    devServer: {
        hot: true,
        headers: {
            "Access-Control-Allow-Origin": "*"
        },
        server: "https",
        port: process.env.npm_package_config_dev_server_port || 3000
    }
};

export default config;
