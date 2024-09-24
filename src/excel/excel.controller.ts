import { Controller, Get, Header, Res } from '@nestjs/common';
import { Response } from 'express';
import { ExcelService } from './excel.service';

@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Get('download')
  @Header('Content-Type', 'text/xlsx')
  async downloadExcel(@Res() res: Response) {
    let result = await this.excelService.downloadExcel();
    res.download(result as string);
  }
}
