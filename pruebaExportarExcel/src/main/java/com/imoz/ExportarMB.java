package com.imoz;

import java.io.IOException;
import java.io.OutputStream;

import javax.faces.bean.ManagedBean;
import javax.faces.bean.SessionScoped;
import javax.faces.context.FacesContext;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

@ManagedBean(name = "exportarMB")
@SessionScoped
public class ExportarMB {

	public void generarExcel() throws IOException {

		HttpServletResponse response = (HttpServletResponse) FacesContext.getCurrentInstance().getExternalContext()
				.getResponse();
		response.addHeader("Content-disposition", "attachment; filename=reporteAlumno.xls");

		response.setContentType("application/vnd.ms-excel");
		try {
			HSSFWorkbook wb = new HSSFWorkbook(); // crea libro de excel
			HSSFSheet sheet = wb.createSheet("Alumnos"); // crea hoja

			HSSFRow row1 = sheet.createRow((short) 0); // crea fila1
			HSSFCell a1 = row1.createCell((short) 0); // crea A1
			HSSFCell b1 = row1.createCell((short) 1); // crea B1

			a1.setCellValue("Alumno");
			b1.setCellValue("Nota");

			HSSFCellStyle cellStyle = wb.createCellStyle();
			cellStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			b1.setCellStyle(cellStyle);

			HSSFRow row2 = sheet.createRow((short) 1); // crea fila2
			row2.createCell((short) 0).setCellValue("Juan"); // A2
			row2.createCell((short) 1).setCellValue(16); // B2

			HSSFRow row3 = sheet.createRow((short) 2); // crea fila3
			row3.createCell((short) 0).setCellValue("Ana"); // A3
			row3.createCell((short) 1).setCellValue(14); // B3

			HSSFRow row4 = sheet.createRow((short) 3); // crea fila4
			row4.createCell((short) 0).setCellValue("Luis"); // A4
			row4.createCell((short) 1).setCellValue(18); // B4

			HSSFRow row5 = sheet.createRow((short) 4); // crea fila5
			row5.createCell((short) 0).setCellValue("Promedio"); // A5
			row5.createCell((short) 1).setCellFormula("Average(B2:B4)");

			OutputStream out = response.getOutputStream();
			wb.write(out);

		} catch (Exception ex) {
			System.out.println(ex.getMessage());
		}

		response.getOutputStream().flush();
		FacesContext.getCurrentInstance().responseComplete();

	}

}
