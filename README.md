# Excelizer
Simplfies Excel data-driven test in ***TestNG*** and ***JUnit***
You don't have to build your own Excel reader library nor have to get knowledge how Apache poi works!!

## Download and Installation
Add the dependency from this [link.
](https://mvnrepository.com/artifact/com.github.amado-saladino/excelizer/1.0.1)

### Maven sample
```
<dependency>
  <groupId>com.github.amado-saladino</groupId>
  <artifactId>excelizer</artifactId>
  <version>1.0.1</version>
</dependency>
```
### Gradle Sample
```
implementation 'com.github.amado-saladino:excelizer:1.0.1'
```
### How to use

Use multiple data providers to provide data from more than one ***Excel*** workbook or sheets. 

***Note***:  In this snippet, it is supposed Excel workbook exists under folder **Data** which is in the project root folder.

```
public class Test1 {
	
	@DataProvider(name = "DataProviderForUsers")
	public Object[][] usersData() {		
		return ExcelReader.loadTestData("Data\\TestData.xlsx","register");
	}
	
	@DataProvider(name = "DataProviderForNames")
	public Object[][] userPhone() {
		return ExcelReader.loadTestData("Data\\TestData.xlsx","sheet2");
	}
	
	@Test(dataProvider="DataProviderForUsers")
	public void test1(String name,String email,String password,String gender) {
		
		System.out.println("This is test...");
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
		System.out.println(phone);
		System.out.println("......");
	}
}
```
Good luck :)
