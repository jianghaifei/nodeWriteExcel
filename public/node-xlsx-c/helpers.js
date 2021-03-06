"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

exports.__esModule = true;
exports.buildSheetFromMatrix = exports.buildExcelDate = exports.isCellDescriptor = exports.isObject = exports.isString = exports.isNumber = exports.isBoolean = void 0;

var _xlsx = _interopRequireDefault(require("xlsx"));

var ORIGIN_DATE = new Date(Date.UTC(1899, 11, 30));

var isBoolean = function isBoolean(maybeBoolean) {
    return typeof maybeBoolean === "boolean";
};

exports.isBoolean = isBoolean;

var isNumber = function isNumber(maybeNumber) {
    return typeof maybeNumber === "number";
};

exports.isNumber = isNumber;

var isString = function isString(maybeString) {
    return typeof maybeString === "string";
};

exports.isString = isString;

var isObject = function isObject(maybeObject) {
    return maybeObject !== null && typeof maybeObject === "object";
};

exports.isObject = isObject;

var isCellDescriptor = function isCellDescriptor(maybeCell) {
    return isObject(maybeCell) && "v" in maybeCell;
};

exports.isCellDescriptor = isCellDescriptor;

var buildExcelDate = function buildExcelDate(value, is1904) {
    var epoch = Date.parse(value + (is1904 ? 1462 : 0));
    return (epoch - ORIGIN_DATE) / 864e5;
};

exports.buildExcelDate = buildExcelDate;

var buildSheetFromMatrix = function buildSheetFromMatrix(data, options) {
    if (options === void 0) {
        options = {};
    }

    var workSheet = {};
    var range = {
        s: {
            c: 1e7,
            r: 1e7,
        },
        e: {
            c: 0,
            r: 0,
        },
    };
    if (!Array.isArray(data)) throw new Error("sheet data is not array");

    for (var R = 0; R !== data.length; R += 1) {
        for (var C = 0; C !== data[R].length; C += 1) {
            if (!Array.isArray(data[R]))
                throw new Error(`${R}th row data is not array`);
            if (range.s.r > R) range.s.r = R;
            if (range.s.c > C) range.s.c = C;
            if (range.e.r < R) range.e.r = R;
            if (range.e.c < C) range.e.c = C;

            if (data[R][C] === null) {
                continue; // eslint-disable-line
            }

            var cell = isCellDescriptor(data[R][C])
                ? data[R][C]
                : {
                      v: data[R][C],
                  };

            var cellRef = _xlsx.default.utils.encode_cell({
                c: C,
                r: R,
            });

            if (isNumber(cell.v)) {
                cell.t = "n";
            } else if (isBoolean(cell.v)) {
                cell.t = "b";
            } else if (cell.v instanceof Date) {
                cell.t = "n";
                cell.v = buildExcelDate(cell.v);
                cell.z = cell.z || _xlsx.default.SSF._table[14]; // eslint-disable-line no-underscore-dangle

                /* eslint-disable spaced-comment, no-trailing-spaces */

                /***
                 * Allows for an non-abstracted representation of the data
                 *
                 * example: {t:'n', z:10, f:'=AVERAGE(A:A)'}
                 *
                 * Documentation:
                 * - Cell Object: https://sheetjs.gitbooks.io/docs/#cell-object
                 * - Data Types: https://sheetjs.gitbooks.io/docs/#data-types
                 * - Format: https://sheetjs.gitbooks.io/docs/#number-formats
                 **/

                /* eslint-disable spaced-comment, no-trailing-spaces */
            } else if (isObject(cell.v)) {
                cell.t = cell.v.t;
                cell.f = cell.v.f;
                cell.z = cell.v.z;
            } else {
                cell.t = "s";
            }

            if (isNumber(cell.z)) cell.z = _xlsx.default.SSF._table[cell.z]; // eslint-disable-line no-underscore-dangle

            workSheet[cellRef] = cell;
        }
    }

    if (range.s.c < 1e7) {
        workSheet["!ref"] = _xlsx.default.utils.encode_range(range);
    }

    if (options["!cols"]) {
        workSheet["!cols"] = options["!cols"];
    }

    if (options["!rows"]) {
        workSheet["!rows"] = options["!rows"];
    }

    if (options["!merges"]) {
        workSheet["!merges"] = options["!merges"];
    }

    return workSheet;
};

exports.buildSheetFromMatrix = buildSheetFromMatrix;
//# sourceMappingURL=helpers.js.map
