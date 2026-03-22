const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyPlugin = require('copy-webpack-plugin');
const webpack = require('webpack');

// ---------------------------------------------------------------------------
// The BASE_URL is where the built files will be hosted.
// For development: https://localhost:3000
// For production:  set via env, e.g.  BASE_URL=https://user.github.io/bi-publisher-mac npm run build
// ---------------------------------------------------------------------------
const DEFAULT_DEV_URL = 'https://localhost:3000';

module.exports = (env, argv) => {
  const isDevelopment = argv.mode === 'development';
  const baseUrl = (env && env.BASE_URL) || (isDevelopment ? DEFAULT_DEV_URL : DEFAULT_DEV_URL);

  // Strip trailing slash
  const cleanBaseUrl = baseUrl.replace(/\/+$/, '');

  return {
    mode: argv.mode || 'development',
    entry: {
      taskpane: './src/taskpane/taskpane.js',
      commands: './src/commands/commands.js'
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].bundle.js',
      publicPath: isDevelopment ? '/' : './',
      clean: true
    },
    devtool: isDevelopment ? 'source-map' : false,
    devServer: {
      host: 'localhost',
      port: 3000,
      https: true,
      allowedHosts: 'all',
      hot: true,
      static: {
        directory: path.join(__dirname, 'public'),
        publicPath: '/'
      },
      headers: {
        'Access-Control-Allow-Origin': '*'
      }
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: 'babel-loader',
            options: {
              presets: [['@babel/preset-env', { targets: { browsers: ['last 2 versions'] } }]]
            }
          }
        },
        {
          test: /\.css$/i,
          use: ['style-loader', 'css-loader']
        },
        {
          test: /\.html$/i,
          loader: 'html-loader'
        },
        {
          test: /\.(png|jpg|jpeg|gif|svg)$/i,
          type: 'asset/resource',
          generator: {
            filename: 'images/[name][ext]'
          }
        }
      ]
    },
    resolve: {
      extensions: ['.js', '.css', '.html'],
      alias: {
        '@': path.resolve(__dirname, 'src/')
      }
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './src/taskpane/taskpane.html',
        chunks: ['taskpane'],
        scriptLoading: 'defer'
      }),
      new HtmlWebpackPlugin({
        filename: 'commands.html',
        template: './src/commands/commands.html',
        chunks: ['commands'],
        scriptLoading: 'defer'
      }),
      new CopyPlugin({
        patterns: [
          {
            from: 'public',
            to: '',
            noErrorOnMissing: true,
            globOptions: { ignore: ['**/.DS_Store'] }
          }
        ]
      }),
      // Generate manifest.xml with the correct BASE_URL baked in
      new CopyPlugin({
        patterns: [
          {
            from: 'manifest.xml',
            to: 'manifest.xml',
            transform(content) {
              return content.toString().replace(/https:\/\/localhost:3000/g, cleanBaseUrl);
            }
          }
        ]
      }),
      new webpack.DefinePlugin({
        'process.env.NODE_ENV': JSON.stringify(argv.mode),
        'process.env.BASE_URL': JSON.stringify(cleanBaseUrl)
      })
    ],
    optimization: {
      minimize: !isDevelopment,
      splitChunks: {
        chunks: 'all',
        cacheGroups: {
          vendor: {
            test: /[\\/]node_modules[\\/]/,
            name: 'vendors',
            priority: 10,
            reuseExistingChunk: true
          }
        }
      }
    },
    performance: {
      hints: isDevelopment ? false : 'warning',
      maxEntrypointSize: 1024000,
      maxAssetSize: 1024000
    },
    stats: {
      preset: isDevelopment ? 'minimal' : 'normal',
      colors: true,
      modules: false,
      children: false
    }
  };
};
