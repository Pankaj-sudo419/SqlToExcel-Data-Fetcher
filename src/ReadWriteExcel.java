//package ExcelDemoProject;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ReadWriteExcel {

public void writeFile() throws IOException{
    FileInputStream file = new FileInputStream("C:\\Users\\Pankaj\\Downloads\\ExceptionHandling\\check\\src\\Book1.xlsx");
    XSSFWorkbook workbook = new XSSFWorkbook(file);
    FileOutputStream outputStream = new FileOutputStream("C:\\Users\\Pankaj\\Downloads\\ExceptionHandling\\check\\src\\Book1.xlsx");

        Sheet sheet = workbook.getSheetAt(0);

        Scanner sc = new Scanner(System.in);
        int rowNum = sheet.getLastRowNum();
    System.out.println("Number of times you want to Enter Data");
        int n = sc.nextInt();
        for(int i=0;i<n;i++) {
        System.out.println("Enter values");

        int rno = sc.nextInt();
        String Name = sc.nextLine();

        Row row = sheet.createRow(++rowNum);
        int columnCount = 0;

        Cell cell = row.createCell(columnCount);
        cell.setCellValue(rno);

        cell = row.createCell(++columnCount);
        cell.setCellValue(Name);
    }
        file.close();

        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
}

public List<Student> readFile() {
    List<Student> list = new ArrayList<>();
    try (FileInputStream file = new FileInputStream("C:\\Users\\Pankaj\\Downloads\\ExceptionHandling\\check\\src\\Book1.xlsx")) {
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = sheet.iterator();

        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            Student s = new Student();
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                switch (currentCell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(currentCell.getStringCellValue() + "\t");
                        s.name = currentCell.getStringCellValue();
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(currentCell.getRowIndex() + "\t");
                        s.rno = currentCell.getRowIndex();
                        break;
                }
            }
            list.add(s);
            System.out.println();
        }
        workbook.close();

    } catch (IOException e) {
        e.printStackTrace();
    }
    return list;
}
    public Connection getCon()
    {
        try
        {
            Class.forName("com.mysql.cj.jdbc.Driver");
            Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/demoproject","root","root");
            return con;
        }
        catch(Exception e)
        {
            System.out.print(e);
            return null;
        }
}
    public static void main(String[] args) throws IOException, SQLException {
        // Read from an existing Excel file
        ReadWriteExcel rw = new ReadWriteExcel();
        rw.writeFile();
        // Write to a Excel file
        List<Student> list = rw.readFile();
        Connection con = rw.getCon();
        Statement stmt = con.createStatement();



        for(int i=1;i<list.size();i++) {
            String name = list.get(i).name;
            int rno = list.get(i).rno;
            stmt.executeUpdate("insert into student (rno,name)values ('"+rno+"','"+name+"')");

        }
        System.out.println("Inserted");

    }
    }




