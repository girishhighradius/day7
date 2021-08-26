import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.*;
public class SqlExl {

	private static Connection getCon() 
	{
		Connection con=null;
		try
		{
			Class.forName("com.mysql.jdbc.Driver");  
			con=DriverManager.getConnection(  
					"jdbc:mysql://10.1.13.80:3306/reportify?useSSL=false","hrcAutoUser","Hrc@1234");  
			 return con;
		}
		catch(Exception e)
		{
			return null;
		}
	}
	public static void main(String[] args) throws SQLException {
		// TODO Auto-generated method stub
		Connection con=getCon();
		if(con != null)
		{
			try {
				Statement smt = con.createStatement();
				ResultSet results = smt.executeQuery("SELECT DISTINCT scenario_id  FROM run_details where run_id=15087");
				
				excelWriter(results);
			}
			
			catch(Exception e)
			{
				e.printStackTrace();
				conClose(con);
			}
			 
			 
		}
	}
	
	
	private static void conClose(Connection con) throws SQLException {
		// TODO Auto-generated method stub
		con.close();
	}
	public static void excelWriter(ResultSet results) {
		// TODO Auto-generated method stub
		XSSFWorkbook workbook = new XSSFWorkbook(); 
		
		XSSFSheet sheet = workbook.createSheet("Senario_ids");
		
		try
		{
			int row_num =0;
			
			while(results.next())
			{
				int cel_num =0;
				Row row = sheet.createRow(row_num++);
				Cell cell = row.createCell(cel_num++);
				
				 cell.setCellValue((Integer)results.getInt(1));
						
				
			}
			
			
			
		}
		catch(Exception e)
		{
			e.printStackTrace();
			
		}
		  try
	        {
	            //Write the workbook in file system
	            FileOutputStream out = new FileOutputStream(new File("senario_ids.xlsx"));
	            workbook.write(out);
	            out.close();
	            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
	        } 
	        catch (Exception e) 
	        {
	            e.printStackTrace();
	        }
		
	}

}
