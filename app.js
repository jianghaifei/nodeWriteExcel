const Koa = require('koa');
const app = new Koa();
const views = require('koa-views');
const json = require('koa-json');
const onerror = require('koa-onerror');
const bodyparser = require('koa-bodyparser');
const logger = require('koa-logger');
const fs = require('fs');
const path = require('path');
const morgan = require('koa-morgan');
const static = require('koa-static');
const redexcel = require('./routes/write');

const ENV = app.env;

// error handler
onerror(app);

const staticPath = './public/';

app.use(static(path.join(__dirname, staticPath)));

// middlewares
app.use(
  bodyparser({
    enableTypes: ['json', 'form', 'text'],
  })
);
app.use(json());
app.use(logger());

// create a write stream (in append mode)
const accessLogStream = fs.createWriteStream(
  path.join(__dirname, './logs', 'access.log'),
  {
    flags: 'a',
  }
);

// setup the logger
if (ENV === 'dev' || ENV === 'test') {
  app.use(morgan('dev'));
} else {
  app.use(
    morgan('combined', {
      stream: accessLogStream,
    })
  );
}

app.use(require('koa-static')(__dirname + '/public/'));

app.use(
  views(__dirname + '/views', {
    extension: 'ejs',
  })
);

// logger
app.use(async (ctx, next) => {
  const start = new Date();
  await next();
  const ms = new Date() - start;
  console.log(`${ctx.method} ${ctx.url} - ${ms}ms`);
});

// routes
app.use(redexcel.routes(), redexcel.allowedMethods());

// error-handling
app.on('error', (err, ctx) => {
  console.error('server error', err, ctx);
});

module.exports = app;
