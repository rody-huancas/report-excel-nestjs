import { Injectable, BadRequestException, NotFoundException } from '@nestjs/common';
import { data } from './data';
import { Workbook } from 'exceljs';
import * as tmp from 'tmp';
import { writeFile } from 'fs/promises';

@Injectable()
export class ExcelService {
  async downloadExcel() {
    if (!data) {
      throw new BadRequestException('No hay datos para descargar');
    }

    let rows = [];

    data.forEach((doc) => { rows.push(Object.values(doc)) });

    let book = new Workbook();
    let sheet = book.addWorksheet('sheet1');

    rows.unshift(Object.keys(data[0]));
    sheet.addRows(rows);

    let File = await new Promise((resolve, reject) => {
      tmp.file(
        {
          discardDescriptor: true,
          prefix           : 'MyExcelSheet',
          postfix          : '.xlsx',
          mode             : parseInt('0600', 8),
        },
        async (err, file) => {
          if (err) throw new BadRequestException(err);

          book.xlsx.writeFile(file)
            .then((_) => { resolve(file) })
            .catch((err) => { reject(err) });
        },
      );
    });

    return File;
  }
}
