const router = require('koa-router')();
const path = require('path');

const { writefile, readPerson } = require('../controller/write');

const { SuccessModel, ErrorModel } = require('../model/resModel');

router.get('/', async (ctx, next) => {
  const read = await readPerson();
  console.log(read);
  await ctx.render('write', {
    read,
  });
});
router.post('/write', async (ctx, next) => {
  const writename = await writefile(ctx.request.body);
  ctx.body = new SuccessModel(writename);
});

module.exports = router;
