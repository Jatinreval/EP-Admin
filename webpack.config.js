const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const webpack = require('webpack');
const dotenv = require('dotenv');
const fs = require('fs');

// Load environment variables from .env file
const env = dotenv.config().parsed || {};

// Create a formatted object for DefinePlugin
const envKeys = Object.keys(env).reduce((prev, next) => {
  prev[`process.env.${next}`] = JSON.stringify(env[next].trim());
  return prev;
}, {});

// Add environment variables from actual process.env (for Render.com)
// This ensures environment variables set on Render are included
Object.keys(process.env).forEach(key => {
  if (!envKeys[`process.env.${key}`] && key.startsWith('AZURE_')) {
    envKeys[`process.env.${key}`] = JSON.stringify(process.env[key].trim());
  }
});

// Create runtime-config.js to make env vars available at runtime
const createRuntimeConfig = () => {
  const configPath = path.resolve(__dirname, 'dist', 'runtime-config.js');
  const configContent = `
    window.__env = {
      AZURE_OPENAI_API_KEY: "${process.env.AZURE_OPENAI_API_KEY || env.AZURE_OPENAI_API_KEY || ''}", 
      AZURE_OPENAI_ENDPOINT: "${process.env.AZURE_OPENAI_ENDPOINT || env.AZURE_OPENAI_ENDPOINT || 'https://epmfl.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-02-15-preview'}"
    };
  `;
  
  // Ensure dist directory exists
  if (!fs.existsSync(path.resolve(__dirname, 'dist'))) {
    fs.mkdirSync(path.resolve(__dirname, 'dist'), { recursive: true });
  }
  
  fs.writeFileSync(configPath, configContent);
  console.log('Runtime config file created at:', configPath);
};

// Log the environment variables for debugging
console.log('Environment variables loaded:', Object.keys(envKeys));

module.exports = {
  mode: process.env.NODE_ENV === 'production' ? 'production' : 'development',
  entry: {
    taskpane: './src/taskpane/taskpane.js',
    commands: './src/commands/commands.js',
    'category-commands': './src/commands/category-commands.js'
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].bundle.js',
    clean: true
  },
  devServer: {
    static: {
      directory: path.join(__dirname, 'dist'),
    },
    headers: {
      "Access-Control-Allow-Origin": "*"
    },
    server: {
      type: 'https',
      options: {
        key: path.join(__dirname, 'certs/server.key'),
        cert: path.join(__dirname, 'certs/server.crt'),
        ca: path.join(__dirname, 'certs/ca.crt')
      }
    },
    port: 3000,
    hot: true,
    setupMiddlewares: (middlewares, devServer) => {
      if (!devServer) {
        throw new Error('webpack-dev-server is not defined');
      }
      
      createRuntimeConfig();
      return middlewares;
    }
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: './src/taskpane/taskpane.html',
      filename: 'taskpane.html',
      chunks: ['taskpane']
    }),
    new HtmlWebpackPlugin({
      template: './src/commands/commands.html',
      filename: 'commands.html',
      chunks: ['commands']
    }),
    new webpack.DefinePlugin({
      // Ensure process.env is defined in the browser
      'process.env': JSON.stringify({}),
      ...envKeys
    }),
    {
      apply: (compiler) => {
        compiler.hooks.afterEmit.tap('CreateRuntimeConfig', (compilation) => {
          createRuntimeConfig();
        });
      }
    }
  ],
  module: {
    rules: [
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader']
      },
      {
        test: /\.(png|svg|jpg|jpeg|gif)$/i,
        type: 'asset/resource',
      }
    ]
  }
}; 