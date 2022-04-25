import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;


public class MainExportExcelFromSQL {
	
static	ArrayList<UserModel> userArrList = new ArrayList<UserModel>();
	
	public static void main(String[] args) throws Exception {
		loadData();
		exportToExcel();
	}
	
	public static void loadData() throws Exception {
		Connection connection = null;
		try {
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			String user= "sa";
			String pass ="sa123";
			String url = "jdbc:sqlserver://localhost:1433;databaseName=NV";
			connection = DriverManager.getConnection(url, user, pass);
			System.out.println("Ket noi den CSDL thanh cong");
			
			String sql = "select * from Info";
			Statement st = connection.createStatement();
			ResultSet rs = st.executeQuery(sql);
			while (rs.next()) {
				int id = rs.getInt("id");
				String name = rs.getString(2);
				float salary = rs.getFloat("salary");
				UserModel us = new UserModel(id, name, salary);
				userArrList.add(us);
				
			}
			connection.close();
		} catch (ClassNotFoundException e) {
		
			e.printStackTrace();
		}
		
	}
	
	public static void exportToExcel() {
		try {
			//XSSF Đọc và ghi định dạng file Open Office XML (XLSX – định dạng hỗ trợ của Excel 2007 trở lên).
		        XSSFWorkbook wb = new XSSFWorkbook();
		        XSSFSheet sheet = wb.createSheet("DanhSachNhanVien");//đặt tên cho bảng tính (sheet)
		        XSSFRow row = null;//1 hàng trong bảng tính
		        Cell cell = null;//1 ô trong 1 hàng
		        
		        row = sheet.createRow(3);//tạo 3 cột
		        
//				tạo thêm 1 cột STT (STT ko có trong DB)
//		        cell = row.createCell(0,CellType.STRING);
//		        cell.setCellValue("STT");
		        
		        cell = row.createCell(1,CellType.STRING);
		        cell.setCellValue("ID");
		        
		        cell = row.createCell(2,CellType.STRING);
		        cell.setCellValue("Name");
		        
		        cell = row.createCell(3,CellType.STRING);
		        cell.setCellValue("Salary");
		        
		        //Duyệt từng phần tử có trong userArrList
		        for (int i = 0; i < userArrList.size(); i++) {
		        	
				row=sheet.createRow(4+i);
				

//				cell = row.createCell(0,CellType.NUMERIC);
//		        cell.setCellValue(i+1); 
		        
				//lấy giá trị ID, Name, Salary đưa vào các Cell tương ứng
		        cell = row.createCell(1,CellType.STRING);
		        cell.setCellValue(userArrList.get(i).getId());
		        
		        cell = row.createCell(2,CellType.STRING);
		        cell.setCellValue(userArrList.get(i).getName());
					
		        cell = row.createCell(3,CellType.STRING);
		        cell.setCellValue(userArrList.get(i).getSalary());
				}
		        
		    //tạo file và đường dẫn
		        File f = new File("danhsachnv.xlsx");
		        
		        try {
					FileOutputStream file = new FileOutputStream(f);
					wb.write(file);
					file.close();
					
				} catch (Exception e) {
					// TODO: handle exception
				}
		 
		} catch (Exception e) {
			// TODO: handle exception
		}
		System.out.println("Xuat file excel thanh cong");
	}
}
