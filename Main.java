import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.concurrent.TimeUnit;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Table;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main 
{

	public static void main(String[] args) throws IOException, InterruptedException, EncryptedDocumentException, InvalidFormatException
	{
		String filelocation = PickAFile();
		System.out.println(filelocation);
		String mergedfilelocation = MergeTextFiles(filelocation);
		System.out.println(mergedfilelocation);
		String excelfilelocation = OpenInExcel(mergedfilelocation, filelocation);
	}

	public static String PickAFile()
	{
	    JFileChooser chooser = new JFileChooser();
	    
	    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
	    
	    chooser.setDialogTitle("Please select the folder that contains all the unrefer for the week.");;
	    
	    int returnVal = chooser.showOpenDialog(null);
	    
	    if(returnVal == JFileChooser.APPROVE_OPTION) 
	    	{
	    		System.out.println("You chose this file: " + chooser.getSelectedFile().getName());
	        	return chooser.getSelectedFile().getAbsolutePath();
	    	}
	    else
	    	System.exit(0);
	    	return null;
	}
	
	public static String MergeTextFiles(String filelocation) throws IOException, InterruptedException
	{
		ProcessBuilder builder = new ProcessBuilder("cmd.exe", "/c", "cd " + filelocation + " && copy *txt mergedfiles.txt");
	        
		builder.redirectErrorStream(true);
	        
	    Process p = builder.start();
	    
	    TimeUnit.SECONDS.sleep(3);
	        
	    return filelocation + "\\mergedfiles.txt";
	}
	
	public static String OpenInExcel(String mergedfilelocation, String filelocation) throws IOException
	{
		LinkedList<String[]> text_lines = new LinkedList<>();
	    
		try (BufferedReader br = new BufferedReader(new FileReader(mergedfilelocation))) 
	    {
	        String sCurrentLine;
	        
	        while ((sCurrentLine = br.readLine()) != null) 
	        {
	        	if(sCurrentLine.length()>2) 
	        	{
		        	String[] sParts = new String[2];
		        	String temp1 = sCurrentLine.substring(0, sCurrentLine.indexOf("$"));
		        	String temp2 = sCurrentLine.substring(sCurrentLine.indexOf("$")+1, sCurrentLine.length());
		        	sParts[0] = temp1;
		        	sParts[1] = temp2;
		            text_lines.add(sParts);
	            }              
	        }
	    } catch (IOException e) 
	    {
	        e.printStackTrace();
	    }
		
		String n = JOptionPane.showInputDialog("What would you like to save the file as?", null);
	    String fileName = filelocation + "/" + n + ".xlsx";
	    XSSFWorkbook workbook = new XSSFWorkbook();
	    XSSFSheet sheet = workbook.createSheet("Main Page");
	    int row_num = 0;
	    
	    for(String[] line : text_lines)
	    {
	        Row row = sheet.createRow(row_num++);
	        int cell_num = 0;
	        
	        for(String value : line)
	        {
	            Cell cell = row.createCell(cell_num++);
	            cell.setCellValue(value);
	        }
		
	    }
	    
	    XSSFSheet s = workbook.createSheet("OURS");
	    Sheet c = workbook.getSheetAt(0);
	    
	    
	    ArrayList<String> Ours = new ArrayList<String>();
	    double ourCount = 0;
	    double theirCount = 0;
	    
	    for (Row row : c) 
	    {
	      for (Cell cell : row) 
	      {
	    	  if (cell.getStringCellValue().equals(""))
	    	  {
	    		  
	    	  }
	    	  else
	    	  {
	    		  if (cell.getStringCellValue().equals("05"))
	    		  {
	    			  Ours.add(row.getCell(0).getStringCellValue());
	    			  //Ours.add(cell.getStringCellValue());
	    			  ourCount = ourCount + 1;
	    		  }
	    		  else
	    		  {
	    			  theirCount = theirCount + 1;
	    		  }
	    	  }
	      }
	    }
	    
	    for(int x = 0; x < Ours.size(); x++)
	    {
	    	System.out.println(Ours.get(x));
	    }
	    
	    theirCount = theirCount / 2 - ourCount/2;
	    
	    double total = theirCount + ourCount;
	    
	    System.out.println(ourCount);
	    System.out.println(theirCount);
	    
	    double ourPercent = (ourCount / total);
	    
	    System.out.println(ourPercent);
	    
	    int counter = 0;
	    
	    row_num = 0;
	    
	    for (String value : Ours)
	    {
	    	Row row = s.createRow(row_num++);
	    	
	    	for (int x = 0; x <1; x++)
	    	{
	    		Cell cell = row.createCell(x);
	            cell.setCellValue(value);
	    	}
	    
	     counter = counter + 1;	
	    }
	    
	    row_num = row_num +3;
	    Row row = s.createRow(row_num);
	    Cell cell = row.createCell(0);
	    cell.setCellValue("Our Percent");
	    row_num++;
	    row = s.createRow(row_num);
	    cell = row.createCell(0);
	    cell.setCellValue(ourPercent);
	    
	    try 
	    {
	        FileOutputStream out = new FileOutputStream(fileName);
	        workbook.write(out);
	        out.close();
	    } 
	    catch (FileNotFoundException ex) 
	    {
	    	System.out.println("Failed.");
	    	System.exit(0);
	    }
	    
	    return fileName;
	
	}
}