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
	 * �ֹ�����һ���򵥸�ʽ��Excel
	 * @return
	 * @throws Exception
	 */
	private static List<Student> getStudent() throws Exception {
		List list = new ArrayList();
		SimpleDateFormat df = new SimpleDateFormat("yyyy-mm-dd");

		Employee e1 = new Employee(1, "����", 16, df.parse("1997-03-12"));
		Employee e2 = new Employee(2, "����", 17, df.parse("1996-08-12"));
		Employee e3 = new Employee(3, "����", 26, df.parse("1958-11-12"));
		list.add(e1);
		list.add(e2);
		list.add(e3);

		return list;
	}

	public static void main(String[] args) throws Exception {
		//��һ������һ��webbook,��Ӧһ��Excel�ļ�
		HSSFWorkbook wb = new HSSFWorkbook();
		//�ڶ�������webbook����� һ��sheet,��Ӧ Excel�ļ��е�sheet
		HSSFSheet sheet = wb.createSheet("ѧ����һ");
		//����������sheet����ӱ�ͷ��0�У�ע���ϰ汾poi��Excel����������������short
		HSSFRow row = sheet.createRow((int) 0);
		//���Ĳ���������Ԫ�񣬲�����ֵ��ͷ�����ñ�ͷ����
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//����һ�����и�ʽ

		HSSFCell cell = row.createCell((short) 0);
		cell.setCellValue("ѧ��");
		cell.setCellStyle(style);
		cell = row.createCell((short) 1);
		cell.setCellValue("����");
		cell.setCellStyle(style);
		cell = row.createCell((short) 2);
		cell.setCellValue("����");
		cell.setCellStyle(style);
		cell = row.createCell((short) 3);
		cell.setCellValue("����");
		cell.setCellStyle(style);

		// ���岽��д��ʵ������ ʵ��Ӧ������Щ���ݴ����ݿ�õ���
		List list = CreateSimpleExcelToDisk.getStudent();

		for (int i = 0; i < list.size(); i++) {
			row = sheet.createRow((int) i + 1);
			Employee stu = (Employee) list.get(i);
			// ���Ĳ���������Ԫ�񣬲�����ֵ
			row.createCell((short) 0).setCellValue((double) stu.getId());
			row.createCell((short) 1).setCellValue(stu.getName());
			row.createCell((short) 2).setCellValue((double) stu.getAge());
			cell = row.createCell((short) 3);
			cell.setCellValue(new SimpleDateFormat("yyyy-mm-dd").format(stu.getBirth()));
		}
		// �����������ļ��浽ָ��λ��
		try {
			FileOutputStream fout = new FileOutputStream("E:/students.xls");
			wb.write(fout);
			fout.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
