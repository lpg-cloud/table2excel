/* eslint-disable camelcase */
/* eslint-disable guard-for-in */
/* eslint-disable no-restricted-syntax */
/* eslint-disable no-shadow */
/* eslint-disable prefer-const */
/* eslint-disable no-redeclare */
/* eslint-disable block-scoped-var */
import {
  saveAs,
} from 'filesaver.js';
import 'xlsx-style/dist/xlsx.core.min';
import dataToWorksheet from './helpers/data-to-worksheet';
import tableToData from './helpers/table-to-data';
import booleanHandler from './types/boolean';
import dateHandler from './types/date';
import inputHandler from './types/input';
import listHandler from './types/list';
import numberHandler from './types/number';


/**
 * @param {string} defaultFileName - The file name if download
 * doesn't provide a name. Default: 'file'.
 * @ param {string} tableNameDataAttribute - The identifier of
 * the name of the table as a data-attribute. Default: 'excel-name'
 * results to `<table data-excel-name="Another table">...</table>`.
 */
const defaultOptions = {
  defaultFileName: 'file',
  tableNameDataAttribute: 'excel-name',
  titleStyle:{
    fill: {
      bgColor: {
        indexed: 64,
      },
      fgColor: {
        rgb: 'FFFF00',
      },
    },
    font: {
      bold: true,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    },
  }
};

/**
 * The default type handlers: lists, numbers, dates, input fields and booleans.
 */
const typeHandlers = [
  listHandler,
  inputHandler,
  numberHandler,
  dateHandler,
  booleanHandler,
];

/**
 * Creates a `Table2Excel` object to export HTMLTableElements
 * to a xlsx-file via its function `export`.
 */
export default class Table2Excel {
  /**
   * @param {object} options - Overrides the default options.
   */
  constructor(options = {}) {
    Object.assign(this, defaultOptions, options);
  }

  /**
   * Exports HTMLTableElements to a xlsx-file.
   *
   * @param {NodeList} tables - The tables to export.
   * @param {string} fileName - The file name.
   */
  export(tables, fileName = this.defaultFileName) {
    this.download(this.getWorkbook(tables), fileName);
  }

  /**
   * Get the XLSX-Workbook object of an array of tables.
   *
   * @param {NodeList} tables - The tables.
   * @returns {object} - The XLSX-Workbook object of the tables.
   */
  getWorkbook(tables) {
    return Array.from(tables.length ? tables : [tables])
      .reduce((workbook, table, index) => {
        const dataName = table.getAttribute(`data-${this.tableNameDataAttribute}`);
        const name = dataName || (index + 1).toString();

        workbook.SheetNames.push(name);
        workbook.Sheets[name] = this.getWorksheet(table);

        return workbook;
      }, {
        SheetNames: [],
        Sheets: {},
      });
  }

  /**
   * Get the XLSX-Worksheet object of a table.
   *
   * @param {HTMLTableElement} table - The table.
   * @returns {object} - The XLSX-Worksheet object of the table.
   */
  getWorksheet(table) {
    if (!table || table.tagName !== 'TABLE') {
      throw new Error('Element must be a table');
    }

    return dataToWorksheet(tableToData(table), typeHandlers);
  }

  /**
   * Exports a XLSX-Workbook object to a xlsx-file.
   *
   * @param {object} workbook - The XLSX-Workbook.
   * @param {string} fileName - The file name.
   */
  download(workbook, fileName = this.defaultFileName) {
    function convert(data) {
      const buffer = new ArrayBuffer(data.length);
      const view = new Uint8Array(buffer);
      for (let i = 0; i <= data.length; i++) {
        view[i] = data.charCodeAt(i) & 0xFF;
      }
      return buffer;
    }

    const data = window.XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'binary',
    });

    const blob = new Blob([convert(data)], {
      type: 'application/octet-stream',
    });
    saveAs(blob, `${fileName}.xlsx`);
  }

  /**
   * 将获取到的tables转换导出到含有一个sheet的workbook中
   * @param {object} tables 一个tables标签[html]对象
   * @param {string} fileName 文件名
   * @param {object} titleStyle 一个cell的style对象
   */
  getWorkbookInOneSheet(tables, fileName = this.defaultFileName, titleStyle = this.titleStyle) {
    let workbook = {
      SheetNames: [fileName],
      Sheets: {},
    };
    let sheet = Array.from(tables.length ? tables : [tables]).reduce((workSheet, table) => {
      let worksheet = this.getWorksheet(table);
      const dataName = table.getAttribute(`data-${this.tableNameDataAttribute}`);

      for (var item in worksheet) {
        if (item[0] != '!') {
          if (worksheet[item].v == '操作' || worksheet[item].v == '详情' || worksheet[item].v == '定位') {
            worksheet[item].v = '';
          }
          // 添加单元格居中
          if (worksheet[item].s) {
            worksheet[item].s.alignment = {
              horizontal: 'center',
              vertical: 'center',
            };
          } else {
            worksheet[item].s = {
              // fill: { bgColor: { indexed: 64 }, fgColor: { rgb: "FFFF00" } },
              alignment: {
                horizontal: 'center',
                vertical: 'center',
              },
            };
          }
        }
      }
      let maxCell=workSheet['!ref'].split(':')[1];
      let maxRowNumb = Number(maxCell.substring(1,maxCell.length)); // 总表格的最大行号
      let maxCelNumnb = maxCell[0]; // 总表个的最大列号

      let currentSheetMaxRowNumber = 0; // 当前表格的最大行号
      let currentSheetMaxCellNumber = 'A'; // 当前表格的最大列号字母
      let currentCellNumber = 0; // 列号数字
      // // 获取总表格的最大行、列
      // for (var cell in workSheet) {
      //   let rowNumber = cell.substring(1, cell.length);
      //   // let rowNumber = workSheet[cell];

      //   if (Number(rowNumber)) {
      //     rowNumber = Number(rowNumber);
      //     if (rowNumber > maxRowNumb) {
      //       maxRowNumb = rowNumber;
      //     }
      //     if (maxCelNumnb.charCodeAt() < cell[0].charCodeAt()) {
      //       maxCelNumnb = cell[0];
      //     }
      //   }
      // }
      // 获取当前表格的最大行、列号
      for (let cell in worksheet) {
        let rowNumber = cell.substring(1, cell.length);
        if (Number(rowNumber)) {
          rowNumber = Number(rowNumber);
          if (rowNumber > currentSheetMaxRowNumber) {
            currentSheetMaxRowNumber = rowNumber;
          }
          if (cell[0].charCodeAt() > currentSheetMaxCellNumber.charCodeAt()) {
            currentSheetMaxCellNumber = cell[0];
            currentCellNumber = currentSheetMaxCellNumber.charCodeAt() - 'A'.charCodeAt();
          }
        }
      }
      if (currentSheetMaxCellNumber.charCodeAt() > maxCelNumnb.charCodeAt()){
        maxCelNumnb = currentSheetMaxCellNumber;
      }
      // 重新设置sheet的ref
      workSheet['!ref'] = 'A1:' + maxCelNumnb + (maxRowNumb + currentSheetMaxRowNumber + 1);

      // 添加表格名称行并设置单元格的样式
      workSheet["A" + (maxRowNumb + 1)] = {
        s: titleStyle,
        t: 'text',
        v: dataName,
      };
      // 将所有的合并单元格添加到总表中
      workSheet['!merges'].push({
        e: {
          r: maxRowNumb,
          c: 0,
        },
        s: {
          r: maxRowNumb,
          c: currentCellNumber - 1,
        },
      });
      // 重新设置当前表格中各个单元格的行号,并将单元格添加到总表中
      for (let cell in worksheet) {
        let rowNumber = cell.substring(1, cell.length);
        if (Number(rowNumber)) {
          rowNumber = Number(rowNumber);
          workSheet[cell[0] + (rowNumber + maxRowNumb + 1)] = worksheet[cell];
        }
      }
      // 重新设置marge,并添加到总表中
      worksheet['!merges'].forEach((merge) => {
        var newMerge = merge;
        newMerge.e.r = merge.e.r + maxRowNumb + 1;
        newMerge.s.r = merge.s.r + maxRowNumb + 1;
        if (merge.e.c == currentCellNumber) {
          // 如果不是合并后两列
          if (merge.s.c != currentCellNumber - 1) {
            newMerge.e.c = merge.e.c - 1;
            workSheet['!merges'].push(newMerge);
          }
        } else {
          workSheet['!merges'].push(newMerge);
        }
      });
      return workSheet;
    }, {
      '!merges': [],
      '!cols': [],
      '!ref': '@0:@0',
    });

    const maxCell = sheet['!ref'].split(':')[1][0];
    for (let i = 0;i< (maxCell.charCodeAt() - 'A'.charCodeAt());i++) {
      sheet['!cols'].push({ wpx: 100 });
    }

    workbook.Sheets[fileName] = sheet;
    this.download(workbook, fileName);
    return workbook;
  }
}

// add global reference to `window` if defined
if (window) window.Table2Excel = Table2Excel;

/**
 * Adds the type handler to the beginning of the list of type handlers.
 * This way it can override general solutions provided by the default handlers
 * with more specific ones.
 *
 * @param {function} typeHandler - Type handler that generates a cell
 * object for a specific cell that fulfills specific criteria.
 * *
 * * @param {HTMLTableCellElement} cell - The cell that should be parsed to a cell object.
 * * @param {string} text - The text of the cell.
 * *
 * * @returns {object} - Cell object (see: https://github.com/SheetJS/js-xlsx#cell-object)
 * * or `null` iff the cell doesn't fulfill the criteria of the type handler.
 */
Table2Excel.extend = function extendCellTypes(typeHandler) {
  typeHandlers.unshift(typeHandler);
};
