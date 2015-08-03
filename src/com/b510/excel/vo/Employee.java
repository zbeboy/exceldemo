package com.b510.excel.vo;

import java.util.Date;

public class Employee {
	private int id;
	private String name;
	private int age;
	private Date birth;

	public Employee() {
	}

	public Employee(int id, String name, int age, Date birth) {
		this.id = id;
		this.name = name;
		this.age = age;
		this.birth = birth;
	}

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public Date getBirth() {
		return birth;
	}

	public void setBirth(Date birth) {
		this.birth = birth;
	}
}
