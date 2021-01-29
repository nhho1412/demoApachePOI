package com.tiger.demoApachePOI.model;

import java.math.BigDecimal;

public class Employee {
	private Integer id;
	private String firstName;
	private String lastName;
	private String address;
	private BigDecimal salary;
	private BigDecimal allowance;
	private BigDecimal totalMoney;

	public Employee() {
		super();
	}

	public Employee(Integer id, String firstName, String lastName, String address, BigDecimal salary, BigDecimal allowance) {
		super();
		this.id = id;
		this.firstName = firstName;
		this.lastName = lastName;
		this.address = address;
		this.salary = salary;
		this.allowance = allowance;
		this.totalMoney = salary.add(allowance);
	}

	@Override
	public String toString() {
		return "Employee[id=" + id + ", firstName=" + firstName + ", lastName=" + lastName + ", address=" + address
				+ ", salary=" + salary + ", allowance=" + allowance + ", totalMoney=" + totalMoney + "]";
	}

	public Integer getId() {
		return id;
	}

	public void setId(Integer id) {
		this.id = id;
	}

	public String getFirstName() {
		return firstName;
	}

	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}

	public String getLastName() {
		return lastName;
	}

	public void setLastName(String lastName) {
		this.lastName = lastName;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}

	public BigDecimal getSalary() {
		return salary;
	}

	public void setSalary(BigDecimal salary) {
		this.salary = salary;
	}

	public BigDecimal getAllowance() {
		return allowance;
	}

	public void setAllowance(BigDecimal allowance) {
		this.allowance = allowance;
	}

	public BigDecimal getTotalMoney() {
		return totalMoney;
	}

	public void setTotalMoney(BigDecimal totalMoney) {
		this.totalMoney = totalMoney;
	}
}
