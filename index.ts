import { Borders, Workbook, Worksheet } from 'exceljs';

enum Movements {
  RIGTH = 'RIGTH',
  LEFT = 'LEFT',
  TOP = 'TOP',
  BOTTOM = 'BOTTOM',
}

enum Actions {
  MOVE = 'MOVE',
  WRITE = 'WRITE',
}

export interface OneSheet {
  name: string;
  data: any;
  headerSchema?: any;
}

export class Complexjson2excelService {
  private sheetConfig = { views: [{ showGridLines: true }] };
  private borderTopStyle: Partial<Borders> = {
    top: {
      color: {
        argb: '000000',
      },
      style: 'thin',
    },
  };
  private lastAction: Actions = Actions.WRITE;
  private lastRow: number = 0;
  private headerFinalRow: number = 0;
  constructor() {}

  getExcelFromJSONList = async (sheets: OneSheet[]) => {
    const workbook = new Workbook();
    workbook.creator = 'Tarjeta Prepaga | Comafi';

    sheets.forEach(({ name, data, headerSchema }) =>
      this.renderSheet(workbook, name, data, headerSchema),
    );

    return await workbook.xlsx.writeBuffer();
  };

  private orderProperties(objList: any[]) {
    if (objList.length === 0) {
      return [];
    }

    const orderedProps = Object.keys(objList[0]);

    function reorderObject(obj: any) {
      const orderedObj = {};
      orderedProps.forEach((prop) => {
        if (typeof obj[prop] !== 'undefined') {
          orderedObj[prop] = obj[prop];
        }
      });
      return orderedObj;
    }

    const orderedList = objList.map((obj) => reorderObject(obj));

    return orderedList;
  }

  private renderSheet = (
    wb: Workbook,
    name: string,
    data: Array<any>,
    headerSchema: any,
  ) => {
    if (!headerSchema && data.length === 0)
      throw new Error(
        'Se necesita un esquema de encabezado o al menos una fila de datos',
      );
    const planned =
      data.length > 0
        ? this.plainArraysToObjects(JSON.parse(JSON.stringify(data[0])))
        : this.plainArraysToObjects(headerSchema);
    data = data.length === 0 ? [] : this.orderProperties(data);
    const worksheet = wb.addWorksheet(name, this.sheetConfig);
    this.renderPyramidHeaders(
      'A1',
      worksheet,
      planned,
      `${name.toUpperCase()} SHEET`,
    );
    this.renderData(worksheet, data);
    this.applyStyles(worksheet);
    this.reset();
  };

  private renderPyramidHeaders = (
    cell: string,
    worksheet: Worksheet,
    data: any,
    name: string,
  ) => {
    const keys = Object.keys(data);
    const baseKeysCount = this.countBaseKeys(data);

    this.write(worksheet, cell, name);
    worksheet.mergeCells(
      `${cell}:${this.move(cell, Movements.RIGTH, baseKeysCount - 1)}`,
    );
    cell = this.moveOne(cell, Movements.BOTTOM);

    for (let i = 0; i < keys.length; i++) {
      if (this.isObject(data[keys[i]])) {
        cell = this.renderPyramidHeaders(
          cell,
          worksheet,
          data[keys[i]],
          keys[i],
        );
        worksheet.getCell(cell).style = {
          font: {
            bold: true,
          },
        };
      } else {
        this.write(worksheet, cell, keys[i]);
        if (i + 1 !== keys.length) cell = this.moveOne(cell, Movements.RIGTH);
      }
    }

    if (this.lastAction === Actions.WRITE)
      cell = this.moveOne(cell, Movements.RIGTH);
    cell = this.moveOne(cell, Movements.TOP);
    this.headerFinalRow = this.lastRow;
    return cell;
  };

  private renderData = (worksheet: Worksheet, data: Array<any>) => {
    const rows = data.map((row) => this.parsePlainArray(row));
    let cell = `A${this.headerFinalRow + 1}`;
    rows.forEach((row) => (cell = this.renderRow(cell, worksheet, row)));
  };

  private renderRow = (
    cell: string,
    worksheet: Worksheet,
    row: Array<any>,
  ): string => {
    const baseRow = this.getNumber(cell);
    let high = 0;

    row.forEach((column) => {
      if (Array.isArray(column)) {
        let highAux = 0;
        worksheet.getCell(cell).style = {
          ...worksheet.getCell(cell).style,
          border: this.borderTopStyle,
        };
        column.forEach((cellValue) => {
          this.write(worksheet, cell, String(cellValue));
          cell = this.moveOne(cell, Movements.BOTTOM);
          highAux++;
        });
        if (highAux > high) high = highAux;
        cell = this.moveOne(
          `${this.getLetter(cell)}${baseRow}`,
          Movements.RIGTH,
        );
      } else {
        this.write(worksheet, cell, String(column));
        worksheet.getCell(cell).style = {
          ...worksheet.getCell(cell).style,
          border: this.borderTopStyle,
        };
        cell = this.moveOne(cell, Movements.RIGTH);
      }
    });

    if (high === 0) high = 1;
    return `A${baseRow + high}`;
  };

  private parsePlainArray = (data) => {
    if (typeof data !== 'object') return data;

    const newArray = [];
    for (const key in data) {
      if (typeof data[key] == 'object' && data[key]) {
        if (Array.isArray(data[key]) && data[key].length > 0) {
          const allObjects = data[key].filter(
            (e) => typeof e == 'object' && !Array.isArray(e),
          );
          if (allObjects.length == data[key].length) {
            if (allObjects.length == 1)
              newArray.push(...this.parsePlainArray(data[key]));
            else {
              const names = Object.keys(allObjects[0]);
              const aux = [];
              for (let index = 0; index < names.length; index++) {
                const pos = names[index];
                const aux2 = data[key].map((e) => {
                  if (typeof e !== 'object') return e[pos];
                  return this.parsePlainArray(e[pos]);
                });
                aux.push(aux2);
              }
              newArray.push(...aux);
            }
          } else {
            newArray.push(
              data[key].map((e) => {
                if (typeof e !== 'object') return e;
                return this.parsePlainArray(e);
              }),
            );
          }
        } else if (Array.isArray(data[key])) newArray.push([]);
        else newArray.push(...this.parsePlainArray(data[key]));
      } else newArray.push(data[key]);
    }
    return newArray;
  };

  private isObject = (data: any) =>
    data && typeof data === 'object' && !Array.isArray(data);

  private countBaseKeys = (object) => {
    let count = 0;

    for (const key in object) {
      if (typeof object[key] === 'object') {
        count += this.countBaseKeys(object[key]);
      } else {
        count++;
      }
    }

    return count;
  };

  private plainArraysToObjects = (object) => {
    for (const key in object) {
      if (Array.isArray(object[key])) {
        if (object[key].length > 0) {
          object[key] = object[key][0];
        } else {
          object[key] = '';
        }
      } else if (typeof object[key] === 'object') {
        this.plainArraysToObjects(object[key]);
      }
    }
    return object;
  };

  private moveOne = (cell: string, movement: Movements) => {
    this.lastAction = Actions.MOVE;
    const number = this.getNumber(cell);
    if (movement === Movements.BOTTOM && number + 1 > this.lastRow)
      this.lastRow = number + 1;
    return this.move(cell, movement);
  };

  private move = (cell: string, movement: Movements, skip = 1) => {
    const number = this.getNumber(cell);
    const letter = this.getLetter(cell);
    switch (movement) {
      case Movements.RIGTH:
        return `${this.nextLetter(letter, skip)}${number}`;
      case Movements.LEFT:
        return `${this.previousLetter(letter, skip)}${number}`;
      case Movements.TOP:
        return `${letter}${number - skip}`;
      case Movements.BOTTOM:
        return `${letter}${number + skip}`;
    }
  };

  private write = (worksheet: Worksheet, cell: string, value: any) => {
    this.lastAction = Actions.WRITE;
    worksheet.getCell(cell).value = value;
  };

  private nextLetter = (letter: string, skip: number = 1): string => {
    let result = '';
    let carry = skip;

    for (let i = letter.length - 1; i >= 0; i--) {
      const char = letter[i];
      const charCode = char.charCodeAt(0) - 65;
      const newCharCode = (charCode + carry) % 26;
      carry = Math.floor((charCode + carry) / 26);

      result = String.fromCharCode(newCharCode + 65) + result;
    }

    if (carry > 0) {
      result = String.fromCharCode(carry + 64) + result;
    }

    return result;
  };

  private previousLetter = (letter: string, skip: number = 1): string => {
    let result = '';
    let borrow = skip;

    for (let i = letter.length - 1; i >= 0; i--) {
      const char = letter[i];
      const charCode = char.charCodeAt(0) - 65;

      if (charCode >= borrow) {
        result = String.fromCharCode(charCode - borrow + 65) + result;
        borrow = 0;
      } else {
        result = String.fromCharCode(charCode + 26 - borrow + 65) + result;
        borrow = 1;
      }
    }

    return result;
  };

  private applyStyles = (worksheet: Worksheet) => {
    worksheet.columns.forEach((column) => {
      let maxLength = 0;
      let lastCellWithValue = null;
      let toMergeCount = 0;
      column['eachCell']({ includeEmpty: true }, (cell) => {
        if (Number(cell.row) <= this.headerFinalRow) {
          if (cell.value && cell.value.toString().length !== 0) {
            if (lastCellWithValue && toMergeCount > 1)
              worksheet.mergeCells(
                `${lastCellWithValue}:${this.move(cell.address, Movements.BOTTOM, toMergeCount)}`,
              );

            lastCellWithValue = cell.address;
            toMergeCount = 0;
          } else toMergeCount++;
          cell.style = {
            ...cell.style,
            font: {
              bold: true,
            },
          };
        }
        cell.style = {
          ...cell.style,
          alignment: {
            horizontal: 'center',
            vertical: 'middle',
          },
        };
        const columnLength = cell.value ? cell.value.toString().length : 10;
        if (columnLength > maxLength) {
          maxLength = columnLength;
        }
      });
      if (toMergeCount > 0 && lastCellWithValue)
        worksheet.mergeCells(
          `${lastCellWithValue}:${this.move(lastCellWithValue, Movements.BOTTOM, toMergeCount)}`,
        );
      column.width = maxLength < 10 ? 10 : maxLength;
    });
  };

  private getNumber = (cell: string) =>
    parseInt(cell.substring(cell.search(/\d/)));

  private getLetter = (cell: string) => cell.substring(0, cell.search(/\d/));

  private reset = () => {
    this.lastAction = Actions.WRITE;
    this.lastRow = 0;
    this.headerFinalRow = 0;
  };
}
