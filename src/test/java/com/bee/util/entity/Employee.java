package com.bee.util.entity;

import java.text.SimpleDateFormat;
import java.util.Date;

public class Employee {
	private Long id;
	private String name;
	private Boolean sex;
	private Integer age;
	private String job;
	private double salery;
	private Date addtime;

	public Long getId() {
		return id;
	}

	public void setId(Long id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Boolean getSex() {
		return sex;
	}

	public void setSex(Boolean sex) {
		this.sex = sex;
	}

	public Integer getAge() {
		return age;
	}

	public void setAge(Integer age) {
		this.age = age;
	}

	public String getJob() {
		return job;
	}

	public void setJob(String job) {
		this.job = job;
	}

	public double getSalery() {
		return salery;
	}

	public void setSalery(double salery) {
		this.salery = salery;
	}

	public Date getAddtime() {
		return addtime;
	}

	public void setAddtime(Date addtime) {
		this.addtime = addtime;
	}

	@Override
	public String toString() {
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String date = null;
		if (addtime != null) {
			date = dateFormat.format(addtime);
		}
		return "Employee [id=" + id + ", name=" + name + ", sex=" + sex + ", age=" + age + ", job=" + job + ", salery=" + salery + ", addtime=" + date + "]";
	}

}
