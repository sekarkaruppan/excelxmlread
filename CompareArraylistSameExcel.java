package working.excell;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


//Compare two excel and write to one excel
public class CompareArraylistSameExcel {
	
    static Boolean check = false;

    public static void main(String args[]) throws IOException {

        try {
        	  List < Employee > excellist1 =    new ArrayList<Employee>();
        	  List < Employee > excellist2 =    new ArrayList<Employee>();
        	  List < Employee > excellist3 =    new ArrayList<Employee>();

            FileInputStream file1 = new FileInputStream(new File(
                    "SameExcel.xlsx"));

           // FileInputStream file2 = new FileInputStream(new File("al2.xlsx"));

            // Get the workbook instance for XLS file
            XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
          //  XSSFWorkbook workbook2 = new XSSFWorkbook(file2);

            // Get first sheet from the workbook
            XSSFSheet sheet1 = workbook1.getSheetAt(0);
            XSSFSheet sheet2 = workbook1.getSheetAt(1);

            // Compare sheets

            // Get iterator to all the rows in current sheet1
            Iterator<Row> rowIterator1 = sheet1.iterator();
            Iterator<Row> rowIterator2 = sheet2.iterator();
            Employee emp;//=new Employee();
            int j=0;
			while (rowIterator1.hasNext()) {
				Row row = rowIterator1.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				emp = new Employee();

				while (cellIterator.hasNext()) {

					Cell cell = cellIterator.next();

					// This is for read only one column from excel
					for (int ii = 0; ii < 3; ii++) {

						if (cell.getColumnIndex() == ii) {
							// Check the cell type and format accordingly
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_NUMERIC:
								System.out.print(cell.getNumericCellValue());
								if (ii == 1)
									emp.setAge(cell.getNumericCellValue()+"");

								break;
							case Cell.CELL_TYPE_STRING:
								System.out.print(cell.getStringCellValue());
								if (ii == 0)
									emp.setId(cell.getStringCellValue().trim());
								if (ii == 2)
									emp.setName(cell.getStringCellValue());
								if (ii == 1)
									emp.setAge(cell.getStringCellValue());
								break;
							case Cell.CELL_TYPE_BOOLEAN:
								System.out.print(cell.getBooleanCellValue());
								break;
							}
						}
					}

				}
				excellist1.add(emp);
				// System.out.println("................ ");
			}

          
			

			while (rowIterator2.hasNext()) {
				Row row = rowIterator2.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				emp = new Employee();

				while (cellIterator.hasNext()) {

					Cell cell = cellIterator.next();

					// This is for read only one column from excel
					for (int ii = 0; ii < 2; ii++) {

						if (cell.getColumnIndex() == ii) {
							// Check the cell type and format accordingly
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_NUMERIC:
								System.out.print(cell.getNumericCellValue());
								
								if (ii == 1)
									emp.setSalary(cell.getNumericCellValue()+"");
								break;
							case Cell.CELL_TYPE_STRING:
								System.out.print(cell.getStringCellValue());
								if (ii == 0)
									emp.setId(cell.getStringCellValue().trim());
								if (ii == 2)
									emp.setName(cell.getStringCellValue());
								if (ii == 1)
									emp.setSalary(cell.getStringCellValue());
								break;
							case Cell.CELL_TYPE_BOOLEAN:
								System.out.print(cell.getBooleanCellValue());
								break;
							}
						}
					}

				}
				excellist2.add(emp);
				
				System.out.println("................ "+emp.getSalary()+".........."+emp.getId());
			}

            
         System.out.println("book1.xls -- " + excellist1.size());
            System.out.println("book1.xls -- " + excellist2.size());


            writeStudentsListToExcel(excellist1,excellist2,excellist3,workbook1);

            // closing the files
            file1.close();
        //    file2.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

 
	//compare two arraylist and write to one arraylist and  write into new file excel

    private static void writeStudentsListToExcel(List<Employee> excellist1,List<Employee> excellist2,List<Employee> excellist3, XSSFWorkbook workBook) {
    	
    	
    	

        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream("SameExcel.xlsx");

            
            
           // ArrayList<String> missingNumbers = new ArrayList<>();
            Employee list3emp;
            for (int x = 0; x < excellist1.size(); x++) {
            	list3emp=new Employee();
                Employee currentUser = (Employee) excellist1.get(x);
                for (int y = 0; y < excellist2.size(); y++) {


                	  Employee list2 = (Employee) excellist2.get(y);
                	//  System.out.println("   ids "+ list2.getId()+"........"+currentUser.getId());
                    if (currentUser.getId().equals(list2.getId())) {
                    	System.out.println("ddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddd");
                    	list3emp.setId(currentUser.getId());
                    	list3emp.setAge(currentUser.getAge());
                    	list3emp.setName(currentUser.getName());
                    	list3emp.setSalary(list2.getSalary());
                    	 System.out.println("  nnnnnnnn "+ list2.getSalary());
                    	 break;
                    }
                       
                    }
                excellist3.add(list3emp);
                }
            System.out.println("final.xls -- " + excellist3.size());
      
            for (int i = 0; i < excellist3.size(); i++) {
    			System.out.println(excellist3.get(i).getId()+"  "+excellist3.get(i).getSalary());
    		}
            
          
         //   XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet spreadSheet = workBook.getSheetAt(2);
          //  XSSFSheet spreadSheet = workBook.createSheet("email");
            
            XSSFRow row;
            XSSFCell cell0;
            XSSFCell cell1;
            XSSFCell cell2;
            XSSFCell cell3;
            // System.out.println("array size is :: "+minusArray.size());
            int cellnumber = 0;
            Employee emp;
            for (int i1 = 0; i1 < excellist3.size(); i1++) {
            	 Employee list3 = (Employee) excellist3.get(i1);
            			
                row = spreadSheet.createRow(i1);
                cell0 = row.createCell(0);
                cell1 = row.createCell(1);
                cell2 = row.createCell(2);
                cell3 = row.createCell(3);
                
               // System.out.print(cell.getCellStyle());
                cell0.setCellValue(list3.getName().toString().trim());
                
                cell1.setCellValue(list3.getId().toString().trim());
                cell2.setCellValue(list3.getAge()+"");
                cell3.setCellValue(list3.getSalary()+"");
                
            }
            workBook.write(fos);
            
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        catch (IOException e) {
            e.printStackTrace();
        }

    }

    // end -write into new file
}