import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class java_service
{
OracledataSource ods=new Oracledatasource();

	
   public static void main(String[] args) throws Exception 
   {
      Class.forName("com.mysql.jdbc.Driver");
      Connection connect = DriverManager.getConnection("jdbc:mysql://localhost:3306/mydb","cloutdeskrw","cloutdeskrw");
      Statement statement = connect.createStatement();
      ResultSet resultSet = statement.executeQuery("select * from mydb.flagstorydetails");
    //  @SuppressWarnings("resource")
     
     
      
	XSSFWorkbook workbook = new XSSFWorkbook();
      XSSFSheet spreadsheet = workbook.createSheet("mydb");
      XSSFRow row=spreadsheet.createRow(1);
      XSSFCell cell;
      cell=row.createCell(1);
      cell.setCellValue("userid");
      cell=row.createCell(2);
      cell.setCellValue("flagstoryid");
      cell=row.createCell(3);
      cell.setCellValue("flagstorydetail");
      cell=row.createCell(4);
      cell.setCellValue("create_time");
      int i=2;
      while(resultSet.next())
      {
         row=spreadsheet.createRow(i);
         cell=row.createCell(1);
         cell.setCellValue(resultSet.getInt("userid"));
         cell=row.createCell(2);
         cell.setCellValue(resultSet.getString("flagstoryid"));
         cell=row.createCell(3);
         cell.setCellValue(resultSet.getString("flagstorydetal"));
         cell=row.createCell(4);
         cell.setCellValue(resultSet.getString("create_time"));
         i++;
      }
      FileOutputStream out = new FileOutputStream(
      new File("C://xls//lagstorydetails.xlsx"));
      workbook.write(out);
      out.close();
      System.out.println("lagstorydetails.xlsx updated successfully");
   }
}