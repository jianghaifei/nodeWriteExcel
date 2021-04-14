const xlsx = require('node-xlsx'); // 必须引入node-xlsx
const nodeXlsx = require('../public/node-xlsx-c'); // 引入二次封装好的 xlsx-style xlsx
const fs = require('fs');

const writefile = async (query) => {
  return await new Promise((resolve, reject) => {
    // 表头样式
    const headerStyle = {
      font: {
        name: '宋体',
        bold: true,
        sz: '20',
      },
      alignment: {
        horizontal: 'center',
        vertical: 'center',
      },
    };
    // 月份时间样式
    const juzStyle = {
      font: {
        name: '宋体',
        bold: false,
        sz: '12',
      },
      alignment: {
        horizontal: 'center',
        vertical: 'center',
      },
      border: {
        top: {
          style: 'thin',
          color: '#000',
        },
        bottom: {
          style: 'thin',
          color: '#000',
        },
        right: {
          style: 'thin',
          color: '#000',
        },
      },
    };
    // 时间样式
    const timeStyle = {
      font: {
        name: 'Arial',
        bold: false,
        sz: '8',
      },
      alignment: {
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      border: {
        top: {
          style: 'thin',
          color: '#000',
        },
        bottom: {
          style: 'thin',
          color: '#000',
        },
        right: {
          style: 'thin',
          color: '#000',
        },
      },
    };
    // title样式
    const titleStyle1 = {
      font: {
        name: '宋体',
        bold: false,
        sz: '11',
      },
    };
    const titleStyle2 = {
      font: {
        name: '宋体',
        bold: false,
        sz: '12',
      },
    };
    // 特殊人员样式
    const teshuPerson = {
      font: {
        name: '宋体',
        bold: false,
        sz: '18',
      },
      alignment: {
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
    };

    const year = query.time.split('-')[0]; // 年
    const month = parseInt(query.time.split('-')[1]) - 1; // 月
    const today = `${year}-${month + 2}-1`;
    let days = new Date(year, month + 1, 0).getDate(); // 计算每个月有多少天
    let onMonth = [];
    const nowMonth = `${year}-${month + 1}-1 ~ ${year}-${month + 1}-${days}`; // 计算当前月份数据

    let weekarry = []; // 当前月份每一天的信息
    let colArray = []; // 列宽
    let rowArray = [{ hpx: 15.6 }, { hpx: 15.6 }, { hpx: 15.6 }, { hpx: 15.6 }]; // 行高
    // 合并单元格 从A1 ~ AD2
    let range = []; // 定义合并单元格

    // 计算当前月所有数据，包含节假日计算
    const monthday = () => {
      // 2021年节假日安排
      const festival = [
        [
          {
            name: '元旦',
            state: 0,
            data: '2021-01-01',
          },
          {
            name: '元旦',
            state: 1,
            data: '2021-01-02',
          },
          {
            name: '元旦',
            state: 1,
            data: '2021-01-03',
          },
        ],
        [
          {
            name: '调休',
            state: 3,
            data: '2021-02-07',
          },
          {
            name: '除夕',
            state: 0,
            data: '2021-02-11',
          },
          {
            name: '春节',
            state: 0,
            data: '2021-02-12',
          },
          {
            name: '春节',
            state: 1,
            data: '2021-02-13',
          },
          {
            name: '春节',
            state: 1,
            data: '2021-02-14',
          },
          {
            name: '春节',
            state: 0,
            data: '2021-02-15',
          },
          {
            name: '春节',
            state: 0,
            data: '2021-02-16',
          },
          {
            name: '春节',
            state: 0,
            data: '2021-02-17',
          },
          {
            name: '调休',
            state: 3,
            data: '2021-02-20',
          },
        ],
        [],
        [
          {
            name: '清明节',
            state: 1,
            data: '2021-04-03',
          },
          {
            name: '清明节',
            state: 1,
            data: '2021-04-04',
          },
          {
            name: '清明节',
            state: 0,
            data: '2021-04-05',
          },
          {
            name: '调休',
            state: 3,
            data: '2021-04-25',
          },
        ],
        [
          {
            name: '劳动节',
            state: 1,
            data: '2021-05-01',
          },
          {
            name: '劳动节',
            state: 1,
            data: '2021-05-02',
          },
          {
            name: '劳动节',
            state: 0,
            data: '2021-05-03',
          },
          {
            name: '劳动节',
            state: 0,
            data: '2021-05-04',
          },
          {
            name: '劳动节',
            state: 0,
            data: '2021-05-05',
          },
          {
            name: '调休',
            state: 3,
            data: '2021-05-08',
          },
        ],
        [
          {
            name: '端午节',
            state: 1,
            data: '2021-01-12',
          },
          {
            name: '端午节',
            state: 1,
            data: '2021-01-13',
          },
          {
            name: '端午节',
            state: 0,
            data: '2021-01-14',
          },
        ],
        [],
        [],
        [
          {
            name: '调休',
            state: 3,
            data: '2021-09-18',
          },
          {
            name: '中秋节',
            state: 1,
            data: '2021-09-19',
          },
          {
            name: '中秋节',
            state: 0,
            data: '2021-09-20',
          },
          {
            name: '中秋节',
            state: 0,
            data: '2021-09-21',
          },
          {
            name: '调休',
            state: 3,
            data: '2021-09-26',
          },
        ],
        [
          {
            name: '国庆节',
            state: 0,
            data: '2021-10-01',
          },
          {
            name: '国庆节',
            state: 1,
            data: '2021-10-02',
          },
          {
            name: '国庆节',
            state: 1,
            data: '2021-10-03',
          },
          {
            name: '国庆节',
            state: 0,
            data: '2021-10-04',
          },
          {
            name: '国庆节',
            state: 0,
            data: '2021-10-05',
          },
          {
            name: '国庆节',
            state: 0,
            data: '2021-10-06',
          },
          {
            name: '国庆节',
            state: 0,
            data: '2021-10-07',
          },
          {
            name: '调休',
            state: 3,
            data: '2021-10-09',
          },
        ],
        [],
        [],
      ];
      // state  0节日放假 1节日占用周六日放假 2正常上班 3节假日调休 4正常周六日休息 5出差 6请假
      // 正常日期计算
      for (let i = 0; i < days; i++) {
        onMonth.push({
          v: i + 1,
          s: juzStyle,
        });
        colArray.push({ wch: 3.91 });
        const weeks = new Date(year, month, i + 1).getDay();
        if (weeks == 6 || weeks == 0) {
          weekarry[i] = {
            Days: i + 1,
            Weeks: weeks,
            state: 4,
            text: '周末休息',
          };
        } else {
          weekarry[i] = {
            Days: i + 1,
            Weeks: weeks,
            state: 2,
            text: '正常上班',
          };
        }
      }

      // 节假日计算
      for (let i = 0; i < festival[month].length; i++) {
        const festdata = festival[month][i];
        let x = weekarry[Number(festdata.data.split('-')[2]) - 1];

        x.state = festdata.state;

        if (x.state != 3) {
          x.text = festdata.name;
        } else {
          x.text = '节日调休';
        }
      }
      return weekarry;
    };
    monthday();

    // 初始化 表头单元格合并 A1 ~ AE2
    range.push({
      s: { c: 0, r: 0 },
      e: { c: weekarry.length - 1, r: 1 },
    });

    // 初始化表格数据
    let newdata = [
      [
        {
          v: '考 勤 记 录 表',
          s: headerStyle,
        },
      ],
      [],
      [
        {
          v: '考勤时间',
          s: titleStyle1,
        },
        null,
        {
          v: nowMonth,
          s: titleStyle1,
        },
        null,
        null,
        null,
        null,
        null,
        null,
        {
          v: '制表时间',
          s: titleStyle1,
        },
        null,
        {
          v: today,
          s: titleStyle1,
        },
      ],
      onMonth,
    ];

    // 每天时间随机取数
    const getautoStr = () => {
      const num = parseInt(Math.random() * (60 - 30) + 30);
      let num2 = parseInt(Math.random() * (60 - 0) + 0);
      if (num2 < 10) {
        num2 = '0' + String(num2);
      }
      return `08:${num} 18:${num2}`;
    };

    // 根据人数循环添加表格数据
    query.personName.split('，').forEach((info, J) => {
      let A = [];
      let B = [];
      rowArray.push({ hpx: 15.6 }, { hpx: 79 });

      for (let i = 0; i < weekarry.length; i++) {
        // 每人两行，此为第一行数据
        switch (i) {
          case 0:
            A.push({
              v: `工 号:${J + 1}`,
              s: titleStyle2,
            });
            break;
          case 8:
            A.push({
              v: `姓 名:`,
              s: titleStyle2,
            });
            break;
          case 10:
            A.push({
              v: `${info.split('：')[0]}`,
              s: titleStyle2,
            });
            break;
          case 18:
            A.push({
              v: `部 门:`,
              s: titleStyle2,
            });
            break;
          case 20:
            A.push({
              v: `${info.split('：')[1]}`,
              s: titleStyle2,
            });
            break;
          default:
            A.push(null);
            break;
        }
        // 每人两行，此为第二行数据
        if (info.split('：')[2]) {
          if (i == 0) {
            B.push({
              v: info.split('：')[2],
              s: teshuPerson,
            });
            let lie = 5 + J * 2;
            range.push({
              s: { c: 0, r: lie },
              e: { c: weekarry.length - 1, r: lie },
            }); // 特殊人员单元格合并
          } else {
            B.push({
              v: '',
              s: timeStyle,
            });
          }
        } else {
          switch (weekarry[i].state) {
            case 2:
              B.push({
                v: getautoStr(),
                s: timeStyle,
              });
              break;
            case 3:
              B.push({
                v: getautoStr(),
                s: timeStyle,
              });
              break;

            default:
              B.push({
                v: '',
                s: timeStyle,
              });
              break;
          }
        }
      }
      newdata.push(A, B);
    });

    // 文件名称
    let name = `${query.title} ${month + 1}月打卡记录`;

    // 配置属性，分别为 列宽/行高/单元格合并
    const options = {
      '!cols': colArray,
      '!rows': rowArray,
      '!merges': range,
    };

    // 创建二进制流
    const buffer = nodeXlsx.build([{ name: 'sheet1', data: newdata }], options);

    // 生成文件
    fs.writeFileSync('./public/excelnew/' + name + '.xls', buffer, 'binary');

    // 将人员名单配置到本地文件中
    fs.writeFile(
      './public/info/personInfo.json',
      query.personName,
      function (err) {
        if (err) {
          return console.error(err);
        }
      }
    );

    // 返回给前端文件名称
    resolve(name + '.xls');
  });
};

const readPerson = async (query) => {
  return await new Promise((resolve, reject) => {
    fs.readFile('./public/info/personInfo.json', function (err, data) {
      if (err) {
        return console.error(err);
      }
      resolve(data.toString());
    });
  });
};

module.exports = {
  writefile,
  readPerson,
};
