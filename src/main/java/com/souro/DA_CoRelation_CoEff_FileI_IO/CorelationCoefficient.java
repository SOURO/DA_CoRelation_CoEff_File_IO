package com.souro.DA_CoRelation_CoEff_FileI_IO;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CorelationCoefficient {
	public static void main(String args[]) {

		List<Double> x = new LinkedList<Double>();
		List<Double> y = new LinkedList<Double>();
		
		/* Please mention your details path here */
		String excelFilePath = "D://Souro_Code_Practice/DA_CoRelation_CoEff_File_IO/IO/IO.xlsx";
		Workbook workbook = null;

		FileInputStream inputStream;
		try {
			inputStream = new FileInputStream(new File(excelFilePath));
			workbook = new XSSFWorkbook(inputStream);
			org.apache.poi.ss.usermodel.Sheet input_sheet = workbook
					.getSheetAt(0);
			Iterator<Row> iterator = input_sheet.iterator();

			while (iterator.hasNext()) {
				Row nextRow = iterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						x.add(cell.getNumericCellValue());
					}
					cell = cellIterator.next();
					if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						y.add(cell.getNumericCellValue());
					}
				}
			}

			double sumX = 0.0, sumY = 0.0, meanX, meanY, x_pow2, x_mul_y, x_sub_meanX, y_sub_meanY, x_sub_meanX_pow2, y_sub_meanY_pow2, x_sub_meanX_mul_y_sub_meanY, x_pow2_sum=0.0, x_mul_y_sum=0.0, x_sub_meanX_sum = 0.0, y_sub_meanY_sum = 0.0, x_sub_meanX_pow2_sum = 0.0, y_sub_meanY_pow2_sum = 0.0, x_sub_meanX_mul_y_sub_meanY_sum = 0.0;
			double cov_xy, s_x, s_y, r_xy;

			int i, n;
			n = x.size();

			for (i = 0; i < n; i++) {
				sumX += x.get(i);
				sumY += y.get(i);
			}

			meanX = sumX / n;
			meanY = sumY / n;

			org.apache.poi.ss.usermodel.Sheet output_sheet = workbook
					.createSheet("Output");
			Row row = null;
			Cell cell = null;
			
			row = output_sheet.createRow(0);
			cell = row.createCell(1);
			cell.setCellValue((String) "X");
			cell = row.createCell(2);
			cell.setCellValue((String) "Y");
			cell = row.createCell(3);
			cell.setCellValue((String) "X-Xmean");
			cell = row.createCell(4);
			cell.setCellValue((String) "X^2");
			cell = row.createCell(5);
			cell.setCellValue((String) "X*Y");
			cell = row.createCell(6);
			cell.setCellValue((String) "Y-Ymean");
			cell = row.createCell(7);
			cell.setCellValue((String) "(X-Xmean)^2");
			cell = row.createCell(8);
			cell.setCellValue((String) "(Y-Ymean)^2");
			cell = row.createCell(9);
			cell.setCellValue((String) "(X-Xmean)*(Y-Ymean)");

			for (i = 0; i < n; i++) {
				
				 x_pow2 = x.get(i) * x.get(i); 
				 x_pow2_sum += x_pow2; 
				 x_mul_y = x.get(i) * y.get(i); 
				 x_mul_y_sum += x_mul_y;
				 
				x_sub_meanX = x.get(i) - meanX;
				x_sub_meanX_pow2 = x_sub_meanX * x_sub_meanX;
				y_sub_meanY = y.get(i) - meanY;
				y_sub_meanY_pow2 = y_sub_meanY * y_sub_meanY;
				x_sub_meanX_sum += x_sub_meanX;
				y_sub_meanY_sum += y_sub_meanY;
				x_sub_meanX_pow2_sum += x_sub_meanX_pow2;
				y_sub_meanY_pow2_sum += y_sub_meanY_pow2;
				x_sub_meanX_mul_y_sub_meanY = x_sub_meanX * y_sub_meanY;
				x_sub_meanX_mul_y_sub_meanY_sum += x_sub_meanX_mul_y_sub_meanY;

				row = output_sheet.createRow(i+1);
				cell = row.createCell(1);
				cell.setCellValue((Double) x.get(i));
				cell = row.createCell(2);
				cell.setCellValue((Double) y.get(i));
				cell = row.createCell(3);
				cell.setCellValue((Double) x_pow2);
				cell = row.createCell(4);
				cell.setCellValue((Double) x_mul_y);
				cell = row.createCell(5);
				cell.setCellValue((Double) x_sub_meanX);
				cell = row.createCell(6);
				cell.setCellValue((Double) y_sub_meanY);
				cell = row.createCell(7);
				cell.setCellValue((Double) x_sub_meanX_pow2);
				cell = row.createCell(8);
				cell.setCellValue((Double) y_sub_meanY_pow2);
				cell = row.createCell(9);
				cell.setCellValue((Double) x_sub_meanX_mul_y_sub_meanY);

				/*System.out
						.println(x.get(i) + "  " + y.get(i) 
															 +"  "+ x_pow2
															 +"  " + x_mul_y
															 +"  "
															 + x_sub_meanX
								+ "  " + y_sub_meanY + "  " + x_sub_meanX_pow2
								+ "  " + y_sub_meanY_pow2 + "  "
								+ x_sub_meanX_mul_y_sub_meanY);*/
			}
			row = output_sheet.createRow(i+1);
			cell = row.createCell(0);
			cell.setCellValue((String) "Total");
			cell = row.createCell(1);
			cell.setCellValue((Double) sumX);
			cell = row.createCell(2);
			cell.setCellValue((Double) sumY);
			cell = row.createCell(3);
			cell.setCellValue((Double) x_pow2_sum);
			cell = row.createCell(4);
			cell.setCellValue((Double) x_mul_y_sum);
			cell = row.createCell(5);
			cell.setCellValue((Double) x_sub_meanX_sum);
			cell = row.createCell(6);
			cell.setCellValue((Double) y_sub_meanY_sum);
			cell = row.createCell(7);
			cell.setCellValue((Double) x_sub_meanX_pow2_sum);
			cell = row.createCell(8);
			cell.setCellValue((Double) y_sub_meanY_pow2_sum);
			cell = row.createCell(9);
			cell.setCellValue((Double) x_sub_meanX_mul_y_sub_meanY_sum);
			
			++i;
			row = output_sheet.createRow(i+1);
			cell = row.createCell(1);
			cell.setCellValue((String) "Xmean");
			cell = row.createCell(2);
			cell.setCellValue((String) "Ymean");
			
			++i;
			row = output_sheet.createRow(i+1);
			cell = row.createCell(1);
			cell.setCellValue((Double) meanX);
			cell = row.createCell(2);
			cell.setCellValue((Double) meanY);

			/*System.out.println(sumX + "  " + sumY + "  " + 
															  x_pow2_sum+"  "+
															  x_mul_y_sum +"  "
															  +
															 x_sub_meanX_sum
					+ "  " + y_sub_meanY_sum + "  " + x_sub_meanX_pow2_sum
					+ "  " + y_sub_meanY_pow2_sum + "  "
					+ x_sub_meanX_mul_y_sub_meanY_sum);*/

			cov_xy = x_sub_meanX_mul_y_sub_meanY_sum / (n - 1);
			s_x = Math.sqrt(x_sub_meanX_pow2_sum / (n - 1));
			s_y = Math.sqrt(y_sub_meanY_pow2_sum / (n - 1));
			r_xy = cov_xy / (s_x * s_y);
			
			++i;
			row = output_sheet.createRow(i+1);
			cell = row.createCell(1);
			cell.setCellValue((String) "COV_xy");
			cell = row.createCell(2);
			cell.setCellValue((String) "S_x");
			cell = row.createCell(3);
			cell.setCellValue((String) "S_y");
			cell = row.createCell(4);
			cell.setCellValue((String) "R_xy");
			
			row = output_sheet.createRow(i+2);
			cell = row.createCell(1);
			cell.setCellValue((Double) cov_xy);
			cell = row.createCell(2);
			cell.setCellValue((Double) s_x);
			cell = row.createCell(3);
			cell.setCellValue((Double) s_y);
			cell = row.createCell(4);
			cell.setCellValue((Double) r_xy);

			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			/*System.out.println(cov_xy + "  " + s_x + "  " + s_y + "  " + r_xy);*/

			workbook.close();
			outputStream.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
