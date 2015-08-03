package com.b510.excel.client;

import java.io.IOException;
import java.util.List;

import com.b510.common.Common;
import com.b510.excel.ReadExcel;
import com.b510.excel.vo.Student;

public class Client {
	public static void main(String[] args) throws IOException {
		String excel2003_2007 = Common.STUDENT_INFO_XLS_PATH;
		String excel2010 = Common.STUDENT_INFO_XLSX_PATH;
		List<Student> list = new ReadExcel().readExcel(excel2003_2007);
		if (list != null) {
			for (Student student : list) {
				System.out.println("No. : " + student.getNo() + ", name : " + student.getName() + ",age : "
						+ student.getAge() + ", score : " + student.getScore());
			}
		}
		System.out.println("===============================");
		List<Student> list1 = new ReadExcel().readExcel(excel2010);
		if (list1 != null) {
			for (Student student : list1) {
				System.out.println("No. : " + student.getNo() + ", name : " + student.getName() + ",age : "
						+ student.getAge() + ", score : " + student.getScore());
			}
		}
	}
}
