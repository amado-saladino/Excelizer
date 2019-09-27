package com.data.Excelizer;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.ahmed.excelizer.ExcelReader;

public class Test1 {
	
	@DataProvider(name = "DataProviderForUsers")
	public Object[][] usersData() {
		return ExcelReader.loadTestData("data\\TestData.xlsx", "register");
	}
	
	@DataProvider(name = "DataProviderForNames")
	public Object[][] userPhone() {
		return ExcelReader.loadTestData("Data\\TestData.xlsx","sheet2");
	}
	
	@Test(dataProvider="DataProviderForUsers")
	public void test1(String name,String email,String password,String gender) {		
		System.out.println("This is a test...");
		System.out.println(name);
		System.out.println(email);
		System.out.println(password);
		System.out.println(gender);
		System.out.println("......");
	}
	
	@Test(dataProvider="DataProviderForNames")
	public void test2(String name,String phone) {		
		System.out.println("This is test...");
		System.out.println(name);
		
		if (phone.isEmpty()) System.out.println("The field is empty");
		
		System.out.println(phone);
		System.out.println("......");
	}
}
