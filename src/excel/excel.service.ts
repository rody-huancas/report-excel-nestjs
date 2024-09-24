import { Injectable, BadRequestException, InternalServerErrorException } from '@nestjs/common';
import { Workbook, Worksheet } from 'exceljs';
import * as tmp from 'tmp';

import { data } from './data';

@Injectable()
export class ExcelService {
  async downloadExcel(): Promise<string> {
    if (!data || data.length === 0) {
      throw new BadRequestException('No hay datos para descargar');
    }

    const rows = data.map(doc => Object.values(doc));
    
    const columnNames = ['Nombre', 'Correo Electrónico'];
    rows.unshift(columnNames);

    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.addRows(rows);

    this.applySheetStyles(sheet);

    try {
      const filePath = await this.generateTempFile(workbook);
      return filePath;
    } catch (error) {
      throw new InternalServerErrorException('Error al generar el archivo Excel', error.message);
    }
  }

  private applySheetStyles(sheet: Worksheet): void {
    // Ancho de las columnas
    sheet.columns[0].width = 25;
    sheet.columns[1].width = 30;

    // Estilos para la primera fila (encabezados)
    const headerRow           = sheet.getRow(1);
          headerRow.height    = 30.5;
          headerRow.font      = { size: 11.5, bold: true, color: { argb: 'FFFFFF' } };
          headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    // Aplicar color de fondo solo a las celdas ocupadas
    const headerCells = headerRow.cellCount;
    for (let i = 1; i <= headerCells; i++) {
      const cell = headerRow.getCell(i);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '000000' } };
      cell.border = {
        top   : { style: 'thin', color: { argb: '000000' } },
        left  : { style: 'thin', color: { argb: 'FFFFFF' } },
        bottom: { style: 'thin', color: { argb: '000000' } },
        right : { style: 'thin', color: { argb: 'FFFFFF' } },
      };
    }

    // Aplicar colores alternados a las filas de datos
    for (let i = 2; i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      const fillColor = i % 2 === 0 ? 'F0F0F0' : 'E0E0E0';
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
      });
    }
  }

  private generateTempFile(workbook: Workbook): Promise<string> {
    return new Promise((resolve, reject) => {
      tmp.file(
        {
          discardDescriptor: true,
          prefix           : `excel-${new Date().toISOString().slice(0, 10)}`,
          postfix          : '.xlsx',
          mode             : parseInt('0600', 8),
        },
        async (err, filePath) => {
          if (err) {
            return reject(new InternalServerErrorException('Error al crear el archivo temporal'));
          }

          try {
            await workbook.xlsx.writeFile(filePath);
            resolve(filePath);
          } catch (writeError) {
            reject(new InternalServerErrorException('Error al escribir el archivo Excel', writeError.message));
          }
        }
      );
    });
  }
}
