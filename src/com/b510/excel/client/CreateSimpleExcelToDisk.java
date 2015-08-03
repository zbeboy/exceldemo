package com.b510.excel.client;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.b510.excel.vo.Employee;
import com.b510.excel.vo.Student;

public class CreateSimpleExcelToDisk {
	/**
	 * 手工构建一个简单格式的Excel
	 * @return
	 * @throws Exception
	 */
	private static List<Student> getStudent() throws Exception {
		List list = new ArrayList();
		SimpleDateFormat df = new SimpleDateFormat("yyyy-mm-dd");

		Employee e1 = new Employee(1, "张三", 16, df.parse("1997-03-12"));
		Employee e2 = new Employee(2, "李四", 17, df.parse("1996-08-12"));
		Employee e3 = new Employee(3, "王五", 26, df.parse("1958-11-12"));
		list.add(e1);
		list.add(e2);
		list.add(e3);

		return list;
	}

	public static void main(String[] args) throws Exception {
		//第一步创建一个webbook,对应一个Excel文件
		HSSFWorkbook wb = new HSSFWorkbook();
		//第二步，在webbook中添加 一个sheet,对应 Excel文件中的sheet
		HSSFSheet sheet = wb.createSheet("学生表一");
		//第三步，在sheet中添加表头第0行，注意老版本poi对Excel的行数列数有限制short
		HSSFRow row = sheet.createRow((int) 0);
		//第四步，创建单元格，并设置值表头，设置表头居中
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//创建一个居中格式

		HSSFCell cell = row.createCell((short) 0);
		cell.setCellValue("学号");
		cell.setCellStyle(style);
		cell = row.createCell((short) 1);
		cell.setCellValue("姓名");
		cell.setCellStyle(style);
		cell = row.createCell((short) 2);
		cell.setCellValue("年龄");
		cell.setCellStyle(style);
		cell = row.createCell((short) 3);
		cell.setCellValue("生日");
		cell.setCellStyle(style);

		// 第五步，写入实体数据 实际应用中这些数据从数据库得到，
		List list = CreateSimpleExcelToDisk.getStudent();

		for (int i = 0; i < list.size(); i++) {
			row = sheet.createRow((int) i + 1);
			Employee stu = (Employee) list.get(i);
			// 第四步，创建单元格，并设置值
			row.createCell((short) 0).setCellValue((double) stu.getId());
			row.createCell((short) 1).setCellValue(stu.getName());
			row.createCell((short) 2).setCellValue((double) stu.getAge());
			cell = row.createCell((short) 3);
			cell.setCellValue(new SimpleDateFormat("yyyy-mm-dd").format(stu.getBirth()));
		}
		// 第六步，将文件存到指定位置
		try {
			FileOutputStream fout = new FileOutputStream("E:/students.xls");
			wb.write(fout);
			fout.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
