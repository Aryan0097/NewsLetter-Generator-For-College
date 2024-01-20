import java.awt.Color;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.List;

import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JSeparator;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;
import javax.swing.border.LineBorder;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

import java.awt.Toolkit;
import javax.swing.JComboBox;
import javax.swing.JProgressBar;

public class MainWindow extends JFrame {

	private JPanel contentPane;

	/**
	 * Launch the application.
	 */
	static JFrame rbName;
	
	//for all coulmn name-old
	public static void rbname() 
    {
		try {
			rbName= new JFrame("Select Name");
	        rbName.getContentPane().setLayout(new FlowLayout());
	        JPanel n=new JPanel();
	        int numberRB = getColumnCount();
	        JRadioButton[] rblist=new JRadioButton[numberRB];
	        ButtonGroup bg=new ButtonGroup();
	        n.setLayout(new javax.swing.BoxLayout(n, javax.swing.BoxLayout.Y_AXIS));
	        for(int i=0;i<numberRB;i++)
	        {
	            rblist[i]=new JRadioButton((String) getCellData(0, i));
	            bg.add(rblist[i]);
	            n.add(rblist[i]);
	        }
	        JButton rbn= new JButton("Ok");
	        n.add(rbn);
	        
	        rbName.getContentPane().add(n);
	        rbName.setSize(250, (35*numberRB)+50);
	        rbName.setLocation(370, 310);
	        rbName.setVisible(true);
	        rbn.addActionListener((e) -> {
	            filename(numberRB,rblist);
	            rbName.dispose();
	        });
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			MainWindow mainWindow=null;
			try {
				mainWindow = new MainWindow();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			mainWindow.rbnameexception(e1);
			e1.printStackTrace();
		}
        
    }
	public void rbnameexception(Exception e) {
	    JOptionPane.showMessageDialog(this, "Excel data fetching Error."+e);
	}
	
	//funtion for count column in excle sheet
    public static int getColumnCount()
    {
    	XSSFWorkbook workbook=null;
    	int columnCount=0;
    	try {
    		String ExcelPath=excelpath;
            workbook = new XSSFWorkbook(ExcelPath);
            XSSFSheet sheet=workbook.getSheet(ch);
            
            columnCount = sheet.getRow(0).getPhysicalNumberOfCells();
		}
    	catch (IOException e) {
    		MainWindow mainWindow=null;
			try {
				mainWindow = new MainWindow();
			} catch (FileNotFoundException e1) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			mainWindow.rbnameexception(e);
			e.printStackTrace();
		} finally {
			// TODO: handle finally clause
			try {
				workbook.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
        
        
        return columnCount;
    }

	//function for get data from excle sheet cell
    public static Object getCellData(int i,int j)
    {
    	String ExcelPath=null; 
    	Object value=null;
        XSSFWorkbook workbook=null;
    	try {
    		ExcelPath=excelpath;
//    		System.out.println(ExcelPath);
    		workbook = new XSSFWorkbook(ExcelPath);
            XSSFSheet sheet=workbook.getSheet(ch);
            
            DataFormatter formattar = new DataFormatter();
            
            value = formattar.formatCellValue(sheet.getRow(i).getCell(j));
		} catch (IOException e) {
    		MainWindow mainWindow=null;
			try {
				mainWindow = new MainWindow();
			} catch (FileNotFoundException e1) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			mainWindow.rbnameexception(e);
			e.printStackTrace();
		}  finally {
			// TODO: handle finally clause
			try {
				workbook.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
        return value;
        
    }
    
//    public static void generatefileAll() throws Exception
//    {
//    		
//    		createfolder();
//            for(int i=1;i<getRowCount();i++)
//            {
////                Find_Replace_DOCX(i);
//            }
//        
//    }
//    
//    public static void generatefirst() throws Exception
//    {
//       
//    		createfolder();
//            Find_Replace_DOCX(1);
//        
//    }
    
    public static void resetall()
    {
    	excelpath="";
    	wordpath="";
    	locpath="";
    	selectedfield="";
    	ch = null;
    	chname = -1;
    	ct=0;
    }
    String[] num= {"1.Departmental Activity","2.Industry - Institute Interaction","3.Research & Development (  Ex. SSIP/SOIC/Patent filled/ Hachathon etc.)","4.Student Achievements","5.Faculty Achievements","6.Community Engagement (NSS, NCC, Clubs etc.)","7.Sports","8.Papers Presented/Published ","9.Professional Development (STTP/FDP/Workshop/Conference etc. attended only/)","10.M.E Dissertation (Tabular Form) – INFORMATION TECHNOLOGY DEPARTMENT"};
	JComboBox<String> comboBox = new JComboBox<>(num);
	static JTextField txt = new JTextField();
    static JTextField wtxt = new JTextField();
    static JTextField etxt = new JTextField();
    
    public static void changename(String s)
    {
    	newfile=s;
    }
    
    static JFrame chkF;
    public static void chk() throws IOException
    {
        chkF = new JFrame("Select Columns");
        chkF.getContentPane().setLayout(new FlowLayout());
        JPanel p = new JPanel();
 
        int numberCheckBox = getColumnCount();
        JCheckBox[] checkBoxList = new JCheckBox[numberCheckBox];
        
        p.setLayout(new javax.swing.BoxLayout(p, javax.swing.BoxLayout.Y_AXIS));
        for(int i = 0; i < numberCheckBox; i++) {
            checkBoxList[i] = new JCheckBox((String) getCellData(0, i));
            p.add(checkBoxList[i]);
        }
        
        
        JButton ok= new JButton("Ok");
        p.add(ok);
        
        chkF.getContentPane().add(p);
        chkF.setSize(250, (35*numberCheckBox)+50);
        chkF.setLocation(370, 310);
        chkF.setVisible(true);
        ok.addActionListener((e) -> {
        	checkCHK(numberCheckBox,checkBoxList);
            chkF.dispose();
        });
        
    }
    
      static int[] chkarr = new int[100];
	  static int ct=0;
	  
	  public static void checkCHK(int n,JCheckBox[] checkBoxList)
	  {
		  ct=0;
		  for(int i=0;i<n;i++)
		  {
			  if(checkBoxList[i].isSelected())
			  {
				  chkarr[i]=1;
			  }
			  else
			  {
				  chkarr[i]=0;
			  }
			  
		  }
		  ct=n;
	  }
    
  //funtion for count rows in excle sheet
	  public static int getRowCount()
	    {
		  	
	        String ExcelPath=null;
	        XSSFWorkbook workbook=null;
	        int rowCount=0;
	        try  {
	        	ExcelPath=excelpath;
	        	workbook = new XSSFWorkbook(ExcelPath);
				XSSFSheet sheet=workbook.getSheet(ch);
				rowCount=0;
				int rowIndex = 7;
			      while (true) {
			         Row row = sheet.getRow(rowIndex);
			         if (row == null) {
			        	rowCount=rowIndex;
			            break;
			         }
			         Cell cell = row.getCell(1);
			         if (cell == null || cell.getCellType() == CellType.BLANK) {
			        	 rowCount=rowIndex;
			            break;
			         }
			         rowIndex++;
			      }
				workbook.close();
				
			} catch (IOException e) {
	    		MainWindow mainWindow=null;
				try {
					mainWindow = new MainWindow();
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				mainWindow.rbnameexception(e);
				e.printStackTrace();
			}  finally {
				// TODO: handle finally clause
				try {
					workbook.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
	        return rowCount;
	    }
    
	  public void wordfileecxeption(Exception e) {
		    JOptionPane.showMessageDialog(this, "Word file Writting Error."+e);
		}

		public static void DepartmentalActivity() throws InterruptedException
		{
		  	String facultycoordinator=null;
			String studentcoordinator=null;
			String eventdetail=null;
			String eventtype=null;
			String targetstudent=null;
			String plateform=null;
			String date=null;
			String nostudent=null;
			String remark=null;
			XWPFDocument doc=null;
		  	try {
		  		doc = new XWPFDocument(new FileInputStream(wordpath));
		    	List<XWPFTable> tables = doc.getTables();
		    	
		    	for (XWPFTable table : tables) {
		    	    XWPFTableRow row = table.getRow(0);
		    	    XWPFTableCell cell = row.getCell(1);
		    	    List<XWPFParagraph> paragraphs = cell.getParagraphs();
		    	    for (XWPFParagraph p : paragraphs) {
		                if (p.getText().contains("Departmental Activity")) {                     
		                    for(int j=8;j<getRowCount();j++)
		                    {
		                    	StringBuilder sb=new StringBuilder("Information Technology department ");
		                    	facultycoordinator=(String) getCellData(j,1);
		                    	studentcoordinator=(String) getCellData(j,2);
		                    	eventtype=(String) getCellData(j,3);
		                    	eventdetail=(String) getCellData(j,8);
		                    	targetstudent=(String) getCellData(j,4);
		                    	plateform=(String) getCellData(j,7);
		                    	date=(String) getCellData(j,6);
		                    	nostudent=(String) getCellData(j,5);
		                    	remark=(String)getCellData(j,9);
		                    	if(facultycoordinator!="")
		                    	{
		                    		sb.append("under the coordination of faculty members "+facultycoordinator+" ");
		                    	}
		                    	if(studentcoordinator!="")
		                    	{
		                    		sb.append("and student coordinator "+studentcoordinator+" ");
		                    	}
		                    	if(eventtype!="")
		                    	{
		                    		sb.append("organized "+eventtype+" ");
		                    	}
		                    	if(targetstudent!="")
		                    	{
		                    		sb.append("for "+targetstudent+" students ");
		                    	}
		                    	if(eventdetail!="")
		                    	{
		                    		sb.append("on "+eventdetail+" ");
		                    	}
		                    	
		                    	if(plateform!="")
		                    	{
		                    		sb.append("at "+plateform+" ");
		                    	}
		                    	if(date!="")
		                    	{
		                    		sb.append("on "+date+" ");
		                    	}
		                    	sb.append(".");
		                    	if(nostudent!="")
		                    	{
		                    		sb.append(nostudent+" students participated in the event.");
		                    	}
		                    	if(remark!="" && remark!=null)
		                    	{
		                    		sb.append(" "+remark);
		                    	}
		                    	
//		                        String ch="Information Technology department under the coordination of faculty members "+facultycoordinator+" organized "+eventdetail+" for "+targetstudent+" students on "+plateform+" on "+date+". "+nostudent+" students participated in the event.";
		                        XWPFParagraph newParagraph = cell.addParagraph();
		                        newParagraph.createRun().setText(sb.toString());
//		                        newParagraph.createRun().addBreak();  
		                    }
//		                    XWPFParagraph newParagraph = cell.addParagraph();
//		                    newParagraph.createRun().addBreak();
		                    break;
		                }
		            }
		    	}
			}catch (IOException e) {
	    		MainWindow mainWindow=null;
				try {
					mainWindow = new MainWindow();
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				mainWindow.wordfileecxeption(e);
				e.printStackTrace();
			} finally {
				// TODO: handle finally clause
				try {
					doc.write(new FileOutputStream(wordpath));
			    	doc.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		      	
		}
	  
	  public static void rename()
		{
			String sourceFilePath = wordpath;
	        String destinationDirPath = locpath;
	        String newFileName = newfile;
	        
	        // Create file objects for the source file, the destination directory, and the new file
	        File sourceFile = new File(sourceFilePath);
	        File destinationDir = new File(destinationDirPath);
	        File newFile = new File(destinationDir, newFileName);
	        
	        try {
	            // Create input stream for the source file and output stream for the new file
	            FileInputStream inputStream = new FileInputStream(sourceFile);
	            FileOutputStream outputStream = new FileOutputStream(newFile);
	            
	            // Create a byte array to hold the file contents
	            byte[] buffer = new byte[1024];
	            int length;
	            
	            // Read bytes from the input stream and write them to the output stream
	            while ((length = inputStream.read(buffer)) > 0) {
	                outputStream.write(buffer, 0, length);
	            }
	            
	            // Close the input and output streams
	            inputStream.close();
	            outputStream.close();
	            
	            // Print a message indicating that the file was copied successfully
	            System.out.println("File copied successfully.");
	            
	            // Rename the new file using the renameTo() method
	            boolean success = newFile.renameTo(new File(destinationDir, "document_backup_renamed.docx"));
	            
	            // Check if the file was renamed successfully
	            if (success) {
	                System.out.println("File renamed successfully.");
	            } else {
	                System.out.println("File renaming failed.");
	            }
	        } catch (IOException e) {
	            // Print an error message if there was an exception
	            e.printStackTrace();
	        }
		}
		
		public static void PapersPresentedPublished() 
		{
			String facultyname=null;
			String coauthor=null;
			String titalofpaper=null;
			String scopeofpublication=null;
			String detail=null;
			String remark=null;
			String date=null;
			XWPFDocument doc=null;
			try {
				doc = new XWPFDocument(new FileInputStream(wordpath));
		    	List<XWPFTable> tables = doc.getTables();
		    	
		    	for (XWPFTable table : tables) {
		    	    XWPFTableRow row = table.getRow(7);
		    	    XWPFTableCell cell = row.getCell(1);
		    	    List<XWPFParagraph> paragraphs = cell.getParagraphs();
		    	    for (XWPFParagraph p : paragraphs) {
		                if (p.getText().contains("Papers Presented/Published")) {                     
		                    for(int j=6;j<getRowCount();j++)
		                    {
		                    	StringBuilder sb=new StringBuilder("");
		                    	facultyname=(String) getCellData(j,1);
		                    	coauthor=(String) getCellData(j,3);
		                    	scopeofpublication=(String) getCellData(j,4);
		                    	titalofpaper=(String) getCellData(j,2);
		                    	detail=(String) getCellData(j,5);
		                    	remark=(String) getCellData(j,7);
		                    	date=(String) getCellData(j,6);
		                    	if(facultyname!="");
		                    	{
		                    		sb.append(facultyname+" ");
		                    	}
		                    	if(coauthor!="")
		                    	{
		                    		sb.append("along with "+coauthor+" ");
		                    	}
		                    	if(titalofpaper!="")
		                    	{
		                    		sb.append("published a research paper on the title "+titalofpaper+" ");
		                    	}
		                    	if(scopeofpublication!="")
		                    	{
		                    		sb.append("in "+scopeofpublication+" ");
		                    	}
		                    	if(detail!="")
		                    	{
		                    		sb.append(detail+" ");
		                    	}
		                    	
		                    	if(remark!="" && remark!=null)
		                    	{
		                    		sb.append("on "+remark+" ");
		                    	}
		                    	if(date!="")
		                    	{
		                    		sb.append("on "+date);
		                    	}
		                    	sb.append(".");
//		                        String ch="Information Technology department under the coordination of faculty members "+facultycoordinator+" organized "+eventdetail+" for "+targetstudent+" students on "+plateform+" on "+date+". "+nostudent+" students participated in the event.";
		                        XWPFParagraph newParagraph = cell.addParagraph();
		                        newParagraph.createRun().setText(sb.toString());
//		                        newParagraph.createRun().addBreak();  
		                    }
//		                    XWPFParagraph newParagraph = cell.addParagraph();
//		                    newParagraph.createRun().addBreak();
		                    break;
		                }
		            }
		    	}
				
			} catch (IOException e) {
	    		MainWindow mainWindow=null;
				try {
					mainWindow = new MainWindow();
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				mainWindow.wordfileecxeption(e);
				e.printStackTrace();
			} finally {
				// TODO: handle finally clause
				try {
					doc.write(new FileOutputStream(wordpath));
			    	doc.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			
		}
		
		public void displayGenerationError() {
		    JOptionPane.showMessageDialog(this, "word file error.");
		}

		
		public static void findnumber(JComboBox<String> comboBox) throws InterruptedException 
		{
			switch(comboBox.getSelectedIndex())
			{
			case 0:DepartmentalActivity();
					break;
			case 1:IndustryInstituteInteraction();
					break;
			case 2:ResearchAndDevelopment();
					break;
			case 3:StudentAchievements();
					break;
			case 4:FacultyAchievements(); 
					break;
			case 5:CommunityEngagement();
					break;
			case 6:Sports();
					break;
			case 7:PapersPresentedPublished();
					break;
			case 8:ProfessionalDevelopment();
					break;
			case 9:MEDissertation();
					break;
			}
		}
		
//		public void rd()
//		{
//			JOptionPane.showMessageDialog(this, "Word file Writting Error.");
//		}
		
		public static void ResearchAndDevelopment() 
		{
//			MainWindow mainWindow=null;
//			try {
//				mainWindow = new MainWindow();
//			} catch (FileNotFoundException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			} catch (IOException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}
//			mainWindow.rd();;
//			String facultyname=null;
//			String coauthor=null;
//			String titalofpaper=null;
//			String scopeofpublication=null;
//			String detail=null;
//			String remark=null;
//			String date=null;
//			XWPFDocument doc = new XWPFDocument(new FileInputStream(wordpath));
//	    	List<XWPFTable> tables = doc.getTables();
//	    	
//	    	for (XWPFTable table : tables) {
//	    	    XWPFTableRow row = table.getRow(7);
//	    	    XWPFTableCell cell = row.getCell(1);
//	    	    List<XWPFParagraph> paragraphs = cell.getParagraphs();
//	    	    for (XWPFParagraph p : paragraphs) {
//	                if (p.getText().contains("Research & Development (  Ex. SSIP/SOIC/Patent filled/ Hachathon etc.)")) {                     
//	                    for(int j=8;j<getRowCount();j++)
//	                    {
//	                    	StringBuilder sb=new StringBuilder("");
//	                    	facultyname=(String) getCellData(j,1);
//	                    	coauthor=(String) getCellData(j,3);
//	                    	scopeofpublication=(String) getCellData(j,4);
//	                    	titalofpaper=(String) getCellData(j,2);
//	                    	detail=(String) getCellData(j,5);
//	                    	remark=(String) getCellData(j,7);
//	                    	date=(String) getCellData(j,6);
//	                    	if(facultyname!="");
//	                    	{
//	                    		sb.append(facultyname+" ");
//	                    	}
//	                    	if(coauthor!="")
//	                    	{
//	                    		sb.append("along with "+coauthor+" ");
//	                    	}
//	                    	if(titalofpaper!="")
//	                    	{
//	                    		sb.append("published a research paper on the title "+titalofpaper+" ");
//	                    	}
//	                    	if(scopeofpublication!="")
//	                    	{
//	                    		sb.append("in "+scopeofpublication+" ");
//	                    	}
//	                    	if(detail!="")
//	                    	{
//	                    		sb.append(detail+" ");
//	                    	}
//	                    	
//	                    	if(remark!="")
//	                    	{
//	                    		sb.append("on "+remark+" ");
//	                    	}
//	                    	if(date!="")
//	                    	{
//	                    		sb.append("on "+date);
//	                    	}
//	                    	sb.append(".");
////	                        String ch="Information Technology department under the coordination of faculty members "+facultycoordinator+" organized "+eventdetail+" for "+targetstudent+" students on "+plateform+" on "+date+". "+nostudent+" students participated in the event.";
//	                        XWPFParagraph newParagraph = cell.addParagraph();
//	                        newParagraph.createRun().setText(sb.toString());
////	                        newParagraph.createRun().addBreak();  
//	                    }
////	                    XWPFParagraph newParagraph = cell.addParagraph();
////	                    newParagraph.createRun().addBreak();
//	                    break;
//	                }
//	            }
//	    	}
//	    	doc.write(new FileOutputStream(wordpath));
//	    	doc.close();
		}
		
		public static void StudentAchievements()
		{
			String facultyguide=null;
			String studentname=null;
			String activityname=null;
			String activitydetail=null;
			String semester=null;
			String certificate=null;
			String remark=null;
			XWPFDocument doc=null;
			try {
				doc = new XWPFDocument(new FileInputStream(wordpath));
		    	List<XWPFTable> tables = doc.getTables();
		    	
		    	for (XWPFTable table : tables) {
		    	    XWPFTableRow row = table.getRow(3);
		    	    XWPFTableCell cell = row.getCell(1);
		    	    List<XWPFParagraph> paragraphs = cell.getParagraphs();
		    	    for (XWPFParagraph p : paragraphs) {
		                if (p.getText().contains("Student Achievements")) {                     
		                    for(int j=6;j<getRowCount();j++)
		                    {
		                    	StringBuilder sb=new StringBuilder("");
		                    	studentname=(String) getCellData(j,1);
		                    	facultyguide=(String) getCellData(j,2);
		                    	semester=(String) getCellData(j,3);
		                    	activityname=(String) getCellData(j,4);
		                    	remark=(String) getCellData(j,7);
		                    	certificate=(String) getCellData(j,6);
		                    	activitydetail=(String) getCellData(j,5);
		                    	if(studentname!="")
		                    	{
		                    		sb.append(studentname+", a student from ");
		                    	}
		                    	if(semester!="")
		                    	{
		                    		sb.append("semester "+semester+" of Information Technology Department");
		                    	}
		                    	if(facultyguide!="")
		                    	{
		                    		sb.append(", under the guidence of "+facultyguide+" ");
		                    	}
		                    	if(activityname!="")
		                    	{
		                    		sb.append("_________"+activityname+" ");
		                    	}
		                    	if(activitydetail!="")
		                    	{
		                    		sb.append("_____"+activitydetail+" ");
		                    	}
		                    	
		                    	if(certificate!="")
		                    	{
		                    		sb.append("and got Certificate of "+certificate+".");
		                    	}
		                    	if(remark!="" && remark!=null)
		                    	{
		                    		sb.append(" "+remark);
		                    	}
		                    	sb.append(".");
		                    	
//		                        String ch="Information Technology department under the coordination of faculty members "+facultycoordinator+" organized "+eventdetail+" for "+targetstudent+" students on "+plateform+" on "+date+". "+nostudent+" students participated in the event.";
		                        XWPFParagraph newParagraph = cell.addParagraph();
		                        newParagraph.createRun().setText(sb.toString());
//		                        newParagraph.createRun().addBreak();  
		                    }
//		                    XWPFParagraph newParagraph = cell.addParagraph();
//		                    newParagraph.createRun().addBreak();
		                    break;
		                }
		            }
		    	}
			} catch (IOException e) {
	    		MainWindow mainWindow=null;
				try {
					mainWindow = new MainWindow();
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				mainWindow.wordfileecxeption(e);
				e.printStackTrace();
			} finally {
				// TODO: handle finally clause
				try {
					doc.write(new FileOutputStream(wordpath));
			    	doc.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			

		}
		
		public static void FacultyAchievements()
		{
			String facultyname=null;
			String activityname=null;
			String activitydetail=null;
			String date=null;
			String remark=null;
			XWPFDocument doc=null;
			try {
				doc = new XWPFDocument(new FileInputStream(wordpath));
		    	List<XWPFTable> tables = doc.getTables();
		    	
		    	for (XWPFTable table : tables) {
		    	    XWPFTableRow row = table.getRow(4);
		    	    XWPFTableCell cell = row.getCell(1);
		    	    List<XWPFParagraph> paragraphs = cell.getParagraphs();
		    	    for (XWPFParagraph p : paragraphs) {
		                if (p.getText().contains("Faculty Achievements")) {                     
		                    for(int j=6;j<getRowCount();j++)
		                    {
		                    	StringBuilder sb=new StringBuilder("");
		                    	facultyname=(String) getCellData(j,1);
		                    	date=(String) getCellData(j,4);
		                    	activityname=(String) getCellData(j,2);
		                    	remark=(String) getCellData(j,5);
		                    	activitydetail=(String) getCellData(j,3);
		                    	if(facultyname!="")
		                    	{
		                    		sb.append(facultyname+", from the Information Technology Department ");
		                    	}
		                    	if(activityname!="")
		                    	{
		                    		sb.append("_________"+activityname+" ");
		                    	}
		                    	if(activitydetail!="")
		                    	{
		                    		sb.append("_____"+activitydetail+" ");
		                    	}                    	
		                    	if(date!="")
		                    	{
		                    		sb.append("on "+date+".");
		                    	}
		                    	if(remark!="" && remark!=null)
		                    	{
		                    		sb.append(" "+remark);
		                    	}
		                    	sb.append(".");
		                    	
//		                        String ch="Information Technology department under the coordination of faculty members "+facultycoordinator+" organized "+eventdetail+" for "+targetstudent+" students on "+plateform+" on "+date+". "+nostudent+" students participated in the event.";
		                        XWPFParagraph newParagraph = cell.addParagraph();
		                        newParagraph.createRun().setText(sb.toString());
//		                        newParagraph.createRun().addBreak();  
		                    }
//		                    XWPFParagraph newParagraph = cell.addParagraph();
//		                    newParagraph.createRun().addBreak();
		                    break;
		                }
		            }
		    	}
			} catch (IOException e) {
	    		MainWindow mainWindow=null;
				try {
					mainWindow = new MainWindow();
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				mainWindow.wordfileecxeption(e);
				e.printStackTrace();
			} finally {
				// TODO: handle finally clause
				try {
					doc.write(new FileOutputStream(wordpath));
			    	doc.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		
		public static void CommunityEngagement()
		{
			String facultyguide=null;
			String studentname=null;
			String activityname=null;
			String activitydetail=null;
			String semester=null;
			String certificate=null;
			String remark=null;
			
			XWPFDocument doc=null;
			try {
				doc = new XWPFDocument(new FileInputStream(wordpath));
		    	List<XWPFTable> tables = doc.getTables();
		    	
		    	for (XWPFTable table : tables) {
		    	    XWPFTableRow row = table.getRow(5);
		    	    XWPFTableCell cell = row.getCell(1);
		    	    List<XWPFParagraph> paragraphs = cell.getParagraphs();
		    	    for (XWPFParagraph p : paragraphs) {
		                if (p.getText().contains("Community Engagement (NSS, NCC, Clubs etc.)")) {                     
		                    for(int j=6;j<getRowCount();j++)
		                    {
		                    	StringBuilder sb=new StringBuilder("");
		                    	studentname=(String) getCellData(j,0);
		                    	facultyguide=(String) getCellData(j,1);
		                    	semester=(String) getCellData(j,2);
		                    	activityname=(String) getCellData(j,3);
		                    	remark=(String) getCellData(j,6);
		                    	certificate=(String) getCellData(j,5);
		                    	activitydetail=(String) getCellData(j,4);
		                    	if(studentname!="")
		                    	{
		                    		sb.append(studentname+", a student from ");
		                    	}
		                    	if(semester!="")
		                    	{
		                    		sb.append("semester "+semester+" of Information Technology Department");
		                    	}
//		                    	if(facultyguide!="")
//		                    	{
//		                    		sb.append(", under the guidence of "+facultyguide+" ");
//		                    	}
		                    	if(activityname!="")
		                    	{
		                    		sb.append("_________"+activityname+" ");
		                    	}
		                    	if(activitydetail!="")
		                    	{
		                    		sb.append("_____"+activitydetail+" ");
		                    	}
		                    	
		                    	if(certificate!="")
		                    	{
		                    		sb.append("and got Certificate of "+certificate+".");
		                    	}
		                    	if(remark!="" && remark!=null)
		                    	{
		                    		sb.append(" "+remark);
		                    	}
		                    	sb.append(".");
		                    	
//		                        String ch="Information Technology department under the coordination of faculty members "+facultycoordinator+" organized "+eventdetail+" for "+targetstudent+" students on "+plateform+" on "+date+". "+nostudent+" students participated in the event.";
		                        XWPFParagraph newParagraph = cell.addParagraph();
		                        newParagraph.createRun().setText(sb.toString());
//		                        newParagraph.createRun().addBreak();  
		                    }
//		                    XWPFParagraph newParagraph = cell.addParagraph();
//		                    newParagraph.createRun().addBreak();
		                    break;
		                }
		            }
		    	}
			} catch (IOException e) {
	    		MainWindow mainWindow=null;
				try {
					mainWindow = new MainWindow();
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				mainWindow.wordfileecxeption(e);
				e.printStackTrace();
			} finally {
				// TODO: handle finally clause
				try {
					doc.write(new FileOutputStream(wordpath));
			    	doc.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		
		public static void Sports() throws InterruptedException
		{
			String facultyguide=null;
			String studentname=null;
			String activityname=null;
			String activitydetail=null;
			String semester=null;
			String certificate=null;
			String remark=null;
			XWPFDocument doc=null;
			try {
				doc = new XWPFDocument(new FileInputStream(wordpath));
		    	List<XWPFTable> tables = doc.getTables();
		    	
		    	for (XWPFTable table : tables) {
		    	    XWPFTableRow row = table.getRow(6);
		    	    XWPFTableCell cell = row.getCell(1);
		    	    List<XWPFParagraph> paragraphs = cell.getParagraphs();
		    	    for (XWPFParagraph p : paragraphs) {
		                if (p.getText().contains("Sports")) {                     
		                    for(int j=6;j<getRowCount();j++)
		                    {
		                    	StringBuilder sb=new StringBuilder("");
		                    	studentname=(String) getCellData(j,0);
		                    	facultyguide=(String) getCellData(j,1);
		                    	semester=(String) getCellData(j,2);
		                    	activityname=(String) getCellData(j,3);
		                    	remark=(String) getCellData(j,6);
		                    	certificate=(String) getCellData(j,5);
		                    	activitydetail=(String) getCellData(j,4);
		                    	if(studentname!="")
		                    	{
		                    		sb.append(studentname+", a student from ");
		                    	}
		                    	if(semester!="")
		                    	{
		                    		sb.append("semester "+semester+" of Information Technology Department");
		                    	}
//		                    	if(facultyguide!="")
//		                    	{
//		                    		sb.append(", under the guidence of "+facultyguide+" ");
//		                    	}
		                    	if(activityname!="")
		                    	{
		                    		sb.append("_________"+activityname+" ");
		                    	}
		                    	if(activitydetail!="")
		                    	{
		                    		sb.append("_____"+activitydetail+" ");
		                    	}
		                    	
		                    	if(certificate!="")
		                    	{
		                    		sb.append("and got Certificate of "+certificate+".");
		                    	}
		                    	if(remark!="" && remark!=null)
		                    	{
		                    		sb.append(" "+remark);
		                    	}
		                    	sb.append(".");
		                    	
		                    	
//		                        String ch="Information Technology department under the coordination of faculty members "+facultycoordinator+" organized "+eventdetail+" for "+targetstudent+" students on "+plateform+" on "+date+". "+nostudent+" students participated in the event.";
		                        XWPFParagraph newParagraph = cell.addParagraph();
		                        newParagraph.createRun().setText(sb.toString());
//		                        newParagraph.createRun().addBreak();  
		                    }
//		                    XWPFParagraph newParagraph = cell.addParagraph();
//		                    newParagraph.createRun().addBreak();
		                    break;
		                }
		            }
		    	}
			} catch (IOException e) {
	    		MainWindow mainWindow=null;
				try {
					mainWindow = new MainWindow();
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				mainWindow.wordfileecxeption(e);
				e.printStackTrace();
			} finally {
				// TODO: handle finally clause
				try {
					doc.write(new FileOutputStream(wordpath));
			    	doc.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		
		public static void ProfessionalDevelopment()
		{
			String srno=null;
			String tital=null;
			String nameoffaculty=null;
			String nameofdepartment=null;
			String organizedby=null;
			String duration=null;
			XWPFDocument doc=null;
			try {
				doc = new XWPFDocument(new FileInputStream(wordpath));
		    	List<XWPFTable> tables = doc.getTables();
		    	
		    	for (XWPFTable table : tables) {
		    	    XWPFTableRow row = table.getRow(8);
		    	    XWPFTableCell cell = row.getCell(1);
		    	    List<XWPFParagraph> paragraphs = cell.getParagraphs();
		    	    for (XWPFParagraph p : paragraphs) {
		                if (p.getText().contains("Professional Development (STTP/FDP/Workshop/Conference etc. attended only/)")) { 
		                	List<XWPFTable> tables1 = cell.getTables();
		                	for(XWPFTable table1 : tables1) {
		                		int numColumns = table1.getRow(0).getTableCells().size();
		                		for(int j=6;j<getRowCount();j++)
		                		{
		                    		  XWPFTableRow newRow = table1.createRow();
		                    		  
		                    		  XWPFTableCell cell0 = newRow.getCell(0);
		                    		  srno=(String) getCellData(j,0);
		                    		  XWPFParagraph newParagraph0 = cell0.addParagraph();
		                    		  newParagraph0.setIndentationLeft(1440/10);  // 1440 is the number of twips in 1 inch, divide by 5 to get 0.2 cm
//		                    		  newParagraph0.setIndentationRight(1440/5);
		                    		  XWPFRun run0 = newParagraph0.createRun();
		                    		  run0.setFontSize(11);
		                    		  run0.setFontFamily("Calibri (Body)");
		                    		  run0.setText(srno);
//		  	                          newParagraph0.createRun().setText(srno);
		  	                          XWPFTableCell cell1 = newRow.getCell(1);
		  	                          tital=(String) getCellData(j,3);
		  	                          XWPFParagraph newParagraph1 = cell1.addParagraph();
		  	                          newParagraph1.setIndentationLeft(1440/10);  // 1440 is the number of twips in 1 inch, divide by 5 to get 0.2 cm
//		                    		  newParagraph1.setIndentationRight(1440/5);
		  	                          XWPFRun run1 = newParagraph1.createRun();
		                    		  run1.setFontSize(11);
		                    		  run1.setFontFamily("Calibri (Body)");
		                    		  run1.setText(tital);
//			                          newParagraph1.createRun().setText(tital);
			                          XWPFTableCell cell2 = newRow.getCell(2);
		  	                          nameoffaculty=(String) getCellData(j,1);
		  	                          XWPFParagraph newParagraph2 = cell2.addParagraph();
		  	                          newParagraph2.setIndentationLeft(1440/10);  // 1440 is the number of twips in 1 inch, divide by 5 to get 0.2 cm
//		                    		  newParagraph2.setIndentationRight(1440/5);
		  	                          XWPFRun run2 = newParagraph2.createRun();
		                    		  run2.setFontSize(11);
		                    		  run2.setFontFamily("Calibri (Body)");
		                    		  run2.setText(nameoffaculty);
//			                          newParagraph2.createRun().setText(nameoffaculty);
			                          XWPFTableCell cell3 = newRow.getCell(3);
		  	                          nameofdepartment="IT";
		  	                          XWPFParagraph newParagraph3 = cell3.addParagraph();
		  	                          newParagraph3.setIndentationLeft(1440/10);  // 1440 is the number of twips in 1 inch, divide by 5 to get 0.2 cm
//		                    		  newParagraph3.setIndentationRight(1440/5);
		  	                          XWPFRun run3 = newParagraph3.createRun();
		                    		  run3.setFontSize(11);
		                    		  run3.setFontFamily("Calibri (Body)");
		                    		  run3.setText(nameofdepartment);
//			                          newParagraph3.createRun().setText(nameofdepartment);
			                          XWPFTableCell cell4 = newRow.getCell(4);
		  	                          organizedby=(String) getCellData(j,5);
		  	                          XWPFParagraph newParagraph4 = cell4.addParagraph();
		  	                          newParagraph4.setIndentationLeft(1440/10);  // 1440 is the number of twips in 1 inch, divide by 5 to get 0.2 cm
//		                    		  newParagraph4.setIndentationRight(1440/5);
		  	                          XWPFRun run4 = newParagraph4.createRun();
		                    		  run4.setFontSize(11);
		                    		  run4.setFontFamily("Calibri (Body)");
		                    		  run4.setText(organizedby);
//			                          newParagraph4.createRun().setText(organizedby);
			                          XWPFTableCell cell5 = newRow.getCell(5);
			                          if((String) getCellData(j,6)==(String) getCellData(j,7))
			                          {
			                        	  duration=(String) getCellData(j,6)+" 1 Day ";
			  	                          XWPFParagraph newParagraph5 = cell5.addParagraph();
			  	                          newParagraph5.setIndentationLeft(1440/10);  // 1440 is the number of twips in 1 inch, divide by 5 to get 0.2 cm
//			                    		  newParagraph5.setIndentationRight(1440/5);
			  	                          XWPFRun run5 = newParagraph5.createRun();
			                    		  run5.setFontSize(11);
			                    		  run5.setFontFamily("Calibri (Body)");
			                    		  run5.setText(duration);
//				                          newParagraph5.createRun().setText(duration);
			                          }
			                          else
			                          {
			                        	  duration=(String) getCellData(j,6)+" to "+(String) getCellData(j,7);
			  	                          XWPFParagraph newParagraph5 = cell5.addParagraph();
			  	                          newParagraph5.setIndentationLeft(1440/10);  // 1440 is the number of twips in 1 inch, divide by 5 to get 0.2 cm
//			                    		  newParagraph5.setIndentationRight(1440/5);
			  	                          XWPFRun run5 = newParagraph5.createRun();
			                    		  run5.setFontSize(11);
			                    		  run5.setFontFamily("Calibri (Body)");
			                    		  run5.setText(duration);
//				                          newParagraph5.createRun().setText(duration); 
			                          }
			                          
		                    	  }
		                    }
		                }
		            }
		    	}
			} catch (IOException e) {
	    		MainWindow mainWindow=null;
				try {
					mainWindow = new MainWindow();
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				mainWindow.wordfileecxeption(e);
				e.printStackTrace();
			} finally {
				// TODO: handle finally clause
				try {
					doc.write(new FileOutputStream(wordpath));
			    	doc.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		
		public static void MEDissertation()
		{
			String srno=null;
			String title=null;
			String nameoffaculty=null;
			String nameofstudent=null;
			String erno=null;
			XWPFDocument doc=null;
			try {
				doc = new XWPFDocument(new FileInputStream(wordpath));
		    	List<XWPFTable> tables = doc.getTables();
		    	
		    	for (XWPFTable table : tables) {
		    	    XWPFTableRow row = table.getRow(9);
		    	    XWPFTableCell cell = row.getCell(1);
		    	    List<XWPFParagraph> paragraphs = cell.getParagraphs();
		    	    for (XWPFParagraph p : paragraphs) {
		                if (p.getText().contains("M.E Dissertation (Tabular Form) – INFORMATION TECHNOLOGY DEPARTMENT")) { 
		                	List<XWPFTable> tables1 = cell.getTables();
		                	for(XWPFTable table1 : tables1) {
		                		int numColumns = table1.getRow(0).getTableCells().size();
		                		for(int j=1;j<getRowCount();j++)
		                		{
		                    		  XWPFTableRow newRow = table1.createRow();
		                    		  
		                    		  XWPFTableCell cell0 = newRow.getCell(0);
		                    		  
		                    		  XWPFParagraph newParagraph0 = cell0.addParagraph();
		                    		  newParagraph0.setIndentationLeft(1440/10);  // 1440 is the number of twips in 1 inch, divide by 5 to get 0.2 cm
//		                    		  newParagraph0.setIndentationRight(1440/5);
		                    		  XWPFRun run0 = newParagraph0.createRun();
		                    		  run0.setFontSize(11);
		                    		  run0.setFontFamily("Calibri (Body)");
		                    		  run0.setText(j+" ");
//		  	                          newParagraph0.createRun().setText(srno);
		  	                          XWPFTableCell cell1 = newRow.getCell(1);
		  	                          nameofstudent=(String) getCellData(j,1);
//		  	                          erno=(String) getCellData(j,0);
		  	                          XWPFParagraph newParagraph1 = cell1.addParagraph();
		  	                          newParagraph1.setIndentationLeft(1440/10);  // 1440 is the number of twips in 1 inch, divide by 5 to get 0.2 cm
//		                    		  newParagraph1.setIndentationRight(1440/5);
		  	                          XWPFRun run1 = newParagraph1.createRun();
		                    		  run1.setFontSize(11);
		                    		  run1.setFontFamily("Calibri (Body)");
		                    		  run1.setText(nameofstudent
		                    				  );
//		                    		  run1.setText(nameofstudent+" - "+erno);
//			                          newParagraph1.createRun().setText(tital);
			                          XWPFTableCell cell2 = newRow.getCell(2);
		  	                          title=(String) getCellData(j,2);
		  	                          XWPFParagraph newParagraph2 = cell2.addParagraph();
		  	                          newParagraph2.setIndentationLeft(1440/10);  // 1440 is the number of twips in 1 inch, divide by 5 to get 0.2 cm
//		                    		  newParagraph2.setIndentationRight(1440/5);
		  	                          XWPFRun run2 = newParagraph2.createRun();
		                    		  run2.setFontSize(11);
		                    		  run2.setFontFamily("Calibri (Body)");
		                    		  run2.setText(title);
//			                          newParagraph2.createRun().setText(nameoffaculty);
			                          XWPFTableCell cell3 = newRow.getCell(3);
			                          nameoffaculty=(String) getCellData(j,3);
		  	                          XWPFParagraph newParagraph3 = cell3.addParagraph();
		  	                          newParagraph3.setIndentationLeft(1440/10);  // 1440 is the number of twips in 1 inch, divide by 5 to get 0.2 cm
//		                    		  newParagraph3.setIndentationRight(1440/5);
		  	                          XWPFRun run3 = newParagraph3.createRun();
		                    		  run3.setFontSize(11);
		                    		  run3.setFontFamily("Calibri (Body)");
		                    		  run3.setText(nameoffaculty);
//			                          newParagraph3.createRun().setText(nameofdepartment);
			                          
			                          
		                    	  }
		                    }
		                }
		            }
		    	}
			} catch (IOException e) {
	    		MainWindow mainWindow=null;
				try {
					mainWindow = new MainWindow();
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				mainWindow.wordfileecxeption(e);
				e.printStackTrace();
			} finally {
				// TODO: handle finally clause
				try {
					doc.write(new FileOutputStream(wordpath));
			    	doc.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		

		public static void IndustryInstituteInteraction()
		{
			String facultycoordinator=null;
			String studentcoordinator=null;
			String eventdetail=null;
			String eventtype=null;
			String targetstudent=null;
			String plateform=null;
			String date=null;
			String nostudent=null;
			String remark=null;
			XWPFDocument doc=null;
			try {
				doc = new XWPFDocument(new FileInputStream(wordpath));
		    	List<XWPFTable> tables = doc.getTables();
		    	
		    	for (XWPFTable table : tables) {
		    	    XWPFTableRow row = table.getRow(1);
		    	    XWPFTableCell cell = row.getCell(1);
		    	    List<XWPFParagraph> paragraphs = cell.getParagraphs();
		    	    for (XWPFParagraph p : paragraphs) {
		                if (p.getText().contains("Industry Interaction")) {                     
		                    for(int j=8;j<getRowCount();j++)
		                    {
		                    	StringBuilder sb=new StringBuilder("Information Technology department ");
		                    	facultycoordinator=(String) getCellData(j,1);
		                    	studentcoordinator=(String) getCellData(j,2);
		                    	eventtype=(String) getCellData(j,3);
		                    	eventdetail=(String) getCellData(j,8);
		                    	targetstudent=(String) getCellData(j,4);
		                    	plateform=(String) getCellData(j,7);
		                    	date=(String) getCellData(j,6);
		                    	nostudent=(String) getCellData(j,5);
		                    	remark=(String) getCellData(j,9);
		                    	if(facultycoordinator!="")
		                    	{
		                    		sb.append("under the coordination of faculty members "+facultycoordinator+" ");
		                    	}
		                    	if(studentcoordinator!="")
		                    	{
		                    		sb.append("and student coordinator "+studentcoordinator+" ");
		                    	}
		                    	if(eventtype!="")
		                    	{
		                    		sb.append("organized "+eventtype+" ");
		                    	}
		                    	if(targetstudent!="")
		                    	{
		                    		sb.append("for "+targetstudent+" students ");
		                    	}
		                    	if(eventdetail!="")
		                    	{
		                    		sb.append("on "+eventdetail+" ");
		                    	}
		                    	
		                    	if(plateform!="")
		                    	{
		                    		sb.append("at "+plateform+" ");
		                    	}
		                    	if(date!="")
		                    	{
		                    		sb.append("on "+date+" ");
		                    	}
		                    	sb.append(".");
		                    	if(nostudent!="")
		                    	{
		                    		sb.append(nostudent+" students participated in the event.");
		                    	}
		                    	if(remark!=null && remark!="")
		                    	{
		                    		sb.append(" "+remark);
		                    	}
//		                        String ch="Information Technology department under the coordination of faculty members "+facultycoordinator+" organized "+eventdetail+" for "+targetstudent+" students on "+plateform+" on "+date+". "+nostudent+" students participated in the event.";
		                        XWPFParagraph newParagraph = cell.addParagraph();
		                        newParagraph.createRun().setText(sb.toString());
//		                        newParagraph.createRun().addBreak();  
		                    }
//		                    XWPFParagraph newParagraph = cell.addParagraph();
//		                    newParagraph.createRun().addBreak();
		                    break;
		                }
		            }
		    	}
			} catch (IOException e) {
	    		MainWindow mainWindow=null;
				try {
					mainWindow = new MainWindow();
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				mainWindow.wordfileecxeption(e);
				e.printStackTrace();
			} finally {
				// TODO: handle finally clause
				try {
					doc.write(new FileOutputStream(wordpath));
			    	doc.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
	    	
		}
    static int chname = -1;
    public static void filename(int n,JRadioButton[] rblist)
    {
        
        for(int i=0;i<n;i++)
            {
                if(rblist[i].isSelected())
                {
                    chname=i;
                    break;
                }
            }
    }
    
    //function for find and replace word in word
//    public static void Find_Replace_DOCX(int i) throws Exception{
//    	String Fword;
//        String Rword;
//        String Npath;
//        XWPFDocument doc = new XWPFDocument(OPCPackage.open(wordpath));
//        for (XWPFParagraph p : doc.getParagraphs()) {
//         List<XWPFRun> runs = p.getRuns();
//         if (runs != null) {
//          for (XWPFRun r : runs) {
//           String text = r.getText(0);
//             for(int k=0;k<ct;k++)
//             {
//             	if(chkarr[k]==0)
//             	{
//             		continue;
//             	}
//                 Fword="{";
//                 Fword += (String)getCellData(0,k);
//                 Fword += "}";
//                 Rword=(String)getCellData(i,k);
////                 Npath = locpath;
////                 Npath += Rword;
////                 Npath += ".docx";
//                 if (text != null && text.contains(Fword)) {
//                     text = text.replace(Fword, Rword);//your content
//                     r.setText(text, 0);
//                 
//                 }
//             }
//          }
//         }
//        }
//
//        for (XWPFTable tbl : doc.getTables()) {
//         for (XWPFTableRow row : tbl.getRows()) {
//          for (XWPFTableCell cell : row.getTableCells()) {
//           for (XWPFParagraph p : cell.getParagraphs()) {
//            for (XWPFRun r : p.getRuns()) {
//             String text = r.getText(0);
//             for(int k=0;k<ct;k++)
//             {
//             	if(chkarr[k]==0)
//             	{
//             		continue;
//             	}
//                 Fword="{";
//                 Fword += (String)getCellData(0,k);
//                 Fword += "}";
//                 Rword=(String)getCellData(i,k);
////                 Npath = locpath;
////                 Npath += Rword;
////                 Npath += ".docx";
//                 if (text != null && text.contains(Fword)) {
//                     text = text.replace(Fword, Rword);//your content
//                     r.setText(text, 0);
//                 
//                 }
//             }
//            }
//           }
//          }
//         }
//        }
//        Npath=locpath;
//        Npath += (String)getCellData(i,chname);
//        Npath +=".docx";
//        doc.write(new FileOutputStream(Npath));
//        doc.close();
//    }
    
    static String ch = null;
    static String newfile=null;
    public static void Selectedsheetname(int n,JRadioButton[] rblist)
    {
        
        for(int i=0;i<n;i++)
            {
                if(rblist[i].isSelected())
                {
                    ch=rblist[i].getText();
                    try {
						txt.setText(ch);
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
                    break;
                    
                }
            }
    }
	
	static JFrame f;
	  public static void rb()
	  {
		  try {
			  f = new JFrame("Select Sheet");
		      f.getContentPane().setLayout(new FlowLayout());
		        JPanel p = new JPanel();
		        ButtonGroup bg=new ButtonGroup();
		        p.setLayout(new javax.swing.BoxLayout(p, javax.swing.BoxLayout.Y_AXIS));
		        int ctsheet=countsheet();
		        JRadioButton[] rblist=new JRadioButton[ctsheet];
		      
		        
		        for(int i=0;i<ctsheet;i++)
	            {
	                rblist[i]=new JRadioButton(rbname(i));
	                bg.add(rblist[i]);
	                p.add(rblist[i]);
	            }
		      
		        JButton rbok= new JButton("Ok");
		        p.add(rbok);
//		      p.add(c1);
//		      p.add(c2);
//		        this.add(p);
		      f.getContentPane().add(p);
		      f.setSize(300,(38*(ctsheet))+60);
		      f.setLocation(370, 250);
		      f.setVisible(true);
		      rbok.addActionListener((e) -> {
		    	  Selectedsheetname(ctsheet,rblist);
		            f.dispose();
		        });
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			MainWindow mainWindow=null;
			try {
				mainWindow = new MainWindow();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			mainWindow.rbnameexception(e1);
			e1.printStackTrace();
		}	      
	  }
	  
	  public static int countsheet() throws IOException
	  {
		  int cts;
		  FileInputStream inputStream = new FileInputStream(excelpath);
          Workbook workbook = new XSSFWorkbook(inputStream);
          
          cts=workbook.getNumberOfSheets();
          workbook.close();
          inputStream.close();
		  return cts;
	  }
	  
	  
	  public static String rbname(int i) 
	  {
		  FileInputStream inputStream=null;
		  Workbook workbook=null;
		  String rname=null;
		  try {
			  inputStream = new FileInputStream(excelpath);
	          workbook = new XSSFWorkbook(inputStream);
	         
	          rname=workbook.getSheetName(i);
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			MainWindow mainWindow=null;
			try {
				mainWindow = new MainWindow();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			mainWindow.rbnameexception(e1);
			e1.printStackTrace();
		}
		  finally {
				// TODO: handle finally clause
				try {
					workbook.close();
					inputStream.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		  
          
          
          return rname;
	  }
	static String excelpath="";
	static String wordpath="";
	static String locpath="";
	static String selectedfield="";
	private JTextField txtsheet;
	
	 public static void createfolder()
	 {
	     locpath += "\\NEW_CREATED_FILES\\";
	     File newFolder = new File(locpath);
	     boolean success = newFolder.mkdir();
	        
	 }
	
	public static void main(String[] args) {
		try {
		    for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
		        if ("Metal".equals(info.getName())) {
		            javax.swing.UIManager.setLookAndFeel(info.getClassName());
		            break;
		        }
		    }
		} catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
		    java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
		}
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					MainWindow frame = new MainWindow();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 * @throws IOException 
	 * @throws FileNotFoundException 
	 */
	public MainWindow() throws FileNotFoundException, IOException {
		setIconImage(Toolkit.getDefaultToolkit().getImage("NEWSLETTER_GENERATOR/icon/newsicon.png"));
		setPreferredSize(new Dimension(900, 650));
		setFont(new Font("Dialog", Font.BOLD, 13));
		setTitle("Automatic Newsletter Generator");
		setSize(new Dimension(1000, 700));
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(250, 100, 1000, 700);
		

		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));

		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JMenuBar menuBar = new JMenuBar();
		menuBar.setBounds(0, 0, 1002, 29);
		contentPane.add(menuBar);
		
		JMenu mnNewMenu = new JMenu("File");
		mnNewMenu.setBorder(new LineBorder(new Color(0, 0, 0)));
		mnNewMenu.setFont(new Font("Dubai", Font.BOLD, 13));
		menuBar.add(mnNewMenu);
		
		JMenuItem newfile = new JMenuItem("New");
		newfile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
//				locationout.setText("");
				resetall();
			}
		});
		newfile.setFont(new Font("Dubai", Font.BOLD, 12));
		mnNewMenu.add(newfile);
		
		JMenuItem selectexcelfile = new JMenuItem("Select Excel File");
		selectexcelfile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser Excelchoose = new JFileChooser();
		        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xls", "xlsx");
		        Excelchoose.setFileFilter(filter);
		        Excelchoose.showOpenDialog(null);
		        File exc = Excelchoose.getSelectedFile();
		        excelpath = exc.getAbsolutePath();
//		        System.out.println(exc);
			}
		});
		selectexcelfile.setFont(new Font("Dubai", Font.BOLD, 12));
		mnNewMenu.add(selectexcelfile);
		
		JMenuItem selectwordfile = new JMenuItem("Select Word File");
		selectwordfile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser Wordchoose = new JFileChooser();
		        FileNameExtensionFilter filter = new FileNameExtensionFilter("Word Files", "doc", "docx");
		        Wordchoose.setFileFilter(filter);
		        Wordchoose.showOpenDialog(null);
		        File exc = Wordchoose.getSelectedFile();
		        wordpath = exc.getAbsolutePath();
			}
		});
		selectwordfile.setFont(new Font("Dubai", Font.BOLD, 12));
		mnNewMenu.add(selectwordfile);
		
		JSeparator separator = new JSeparator();
		mnNewMenu.add(separator);
		
		JMenuItem exit = new JMenuItem("Exit");
		exit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.exit(0);
			}
		});
		exit.setFont(new Font("Dubai", Font.BOLD, 12));
		mnNewMenu.add(exit);
		
		JMenu editmenu = new JMenu("Edit");
		editmenu.setBorder(new LineBorder(new Color(0, 0, 0)));
		editmenu.setFont(new Font("Dubai", Font.BOLD, 13));
		menuBar.add(editmenu);
		
		JMenuItem editsheet = new JMenuItem("Edit Selected Sheet");
		editsheet.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				rb();
			}
		});
		editsheet.setFont(new Font("Dubai", Font.BOLD, 12));
		editmenu.add(editsheet);
		
		JMenu runmenu = new JMenu("Run");
		runmenu.setBorder(new LineBorder(new Color(0, 0, 0)));
		runmenu.setFont(new Font("Dubai", Font.BOLD, 13));
		menuBar.add(runmenu);
		
		JMenuItem runall = new JMenuItem("Generate");
		runall.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					findnumber(comboBox);
					JOptionPane.showMessageDialog(MainWindow.this, "Done.");
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					JOptionPane.showMessageDialog(MainWindow.this, "Generation error."+e);
					e1.printStackTrace();
				}
			}
		});
		runall.setFont(new Font("Dubai", Font.BOLD, 12));
		runmenu.add(runall);
		
		JMenu helpmenu = new JMenu("Help");
		helpmenu.setBorder(new LineBorder(new Color(0, 0, 0)));
		helpmenu.setFont(new Font("Dubai", Font.BOLD, 13));
		menuBar.add(helpmenu);
		
		JMenuItem userguide = new JMenuItem("User Guide");
		userguide.setFont(new Font("Dubai", Font.BOLD, 12));
		helpmenu.add(userguide);
		
		JMenuItem feedback = new JMenuItem("Feedback");
		feedback.setFont(new Font("Dubai", Font.BOLD, 12));
		helpmenu.add(feedback);
		
		JSeparator separator_2 = new JSeparator();
		helpmenu.add(separator_2);
		
		JMenuItem about = new JMenuItem("About");
		about.setFont(new Font("Dubai", Font.BOLD, 12));
		helpmenu.add(about);
		
		JLabel lblNewLabel = new JLabel("Select Excel File");
		lblNewLabel.setFont(new Font("Tahoma", Font.BOLD, 20));
		lblNewLabel.setBounds(10, 39, 196, 42);
		contentPane.add(lblNewLabel);
		
		JButton selectexcel = new JButton("Browse");
		selectexcel.setBackground(Color.WHITE);
		selectexcel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser Excelchoose = new JFileChooser();
		        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xls", "xlsx");
		        Excelchoose.setFileFilter(filter);
		        Excelchoose.showOpenDialog(null);
		        File exc = Excelchoose.getSelectedFile();
		        excelpath = exc.getAbsolutePath();
		        etxt.setText(excelpath);
//		        System.out.print(excelpath);
//		        JOptionPane.showMessageDialog(MainWindow.this, "Next:Select Sheet.");
			}
		});
		selectexcel.setFont(new Font("Trebuchet MS", Font.BOLD, 17));
		selectexcel.setBounds(301, 48, 117, 29);
		contentPane.add(selectexcel);
		
		JLabel lblNewLabel_1 = new JLabel("Select Sheet");
		lblNewLabel_1.setFont(new Font("Tahoma", Font.BOLD, 20));
		lblNewLabel_1.setBounds(10, 422, 176, 29);
		contentPane.add(lblNewLabel_1);
		
		JButton selectsheet = new JButton("Select Sheet");
		selectsheet.setBackground(Color.WHITE);
		selectsheet.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				rb();
				
			}
		});
		selectsheet.setFont(new Font("Trebuchet MS", Font.BOLD, 17));
		selectsheet.setBounds(274, 424, 144, 29);
		contentPane.add(selectsheet);
		
		JSeparator separator_3 = new JSeparator();
		separator_3.setBounds(0, 138, 456, 15);
		contentPane.add(separator_3);
		
		JLabel lblNewLabel_2 = new JLabel("Select Word File");
		lblNewLabel_2.setFont(new Font("Tahoma", Font.BOLD, 20));
		lblNewLabel_2.setBounds(10, 163, 196, 35);
		contentPane.add(lblNewLabel_2);
		
		JButton selectword = new JButton("Browse");
		selectword.setBackground(Color.WHITE);
		selectword.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser Wordchoose = new JFileChooser();
		        FileNameExtensionFilter filter = new FileNameExtensionFilter("Word Files", "doc", "docx");
		        Wordchoose.setFileFilter(filter);
		        Wordchoose.showOpenDialog(null);
		        File exc = Wordchoose.getSelectedFile();
		        wordpath = exc.getAbsolutePath();
		        wtxt.setText(wordpath);
//		        JOptionPane.showMessageDialog(MainWindow.this, "Next:Select Location Where You Want to Save All Generated Files.");
			}
		});
		selectword.setFont(new Font("Trebuchet MS", Font.BOLD, 17));
		selectword.setBounds(301, 172, 117, 29);
		contentPane.add(selectword);
		
		JSeparator separator_4 = new JSeparator();
		separator_4.setBounds(0, 262, 456, 15);
		contentPane.add(separator_4);
		
		JSeparator separator_4_1 = new JSeparator();
		separator_4_1.setBounds(0, 383, 456, 15);
		contentPane.add(separator_4_1);
		
		JSeparator separator_4_1_1 = new JSeparator();
		separator_4_1_1.setBounds(0, 520, 456, 15);
		contentPane.add(separator_4_1_1);
		
		JSeparator separator_5 = new JSeparator();
		separator_5.setOrientation(SwingConstants.VERTICAL);
		separator_5.setBounds(466, 26, 15, 606);
		contentPane.add(separator_5);
		
		JButton generateall = new JButton("Generate");
		generateall.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
//					generatefileAll();
//					changename(newname.getText());
//					System.out.println(excelpath);
					findnumber(comboBox);
					JOptionPane.showMessageDialog(MainWindow.this, "Done.");
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					JOptionPane.showMessageDialog(MainWindow.this, "Generation error."+e1);
					e1.printStackTrace();
				}
			}
		});
		generateall.setFont(new Font("Tahoma", Font.BOLD, 20));
		generateall.setBounds(156, 574, 144, 42);
		contentPane.add(generateall);
		
		comboBox.setFont(new Font("Tw Cen MT Condensed", Font.BOLD, 20));
		comboBox.setBounds(85, 300, 300, 29);
		contentPane.add(comboBox);
		
		
		txt.setBounds(10, 474, 300, 19);
		contentPane.add(txt);
		txt.setColumns(10);
		
		etxt.setBounds(10, 90, 450, 19);
		contentPane.add(etxt);
		etxt.setColumns(25);
		
		wtxt.setBounds(10, 220, 450, 19);
		contentPane.add(wtxt);
		wtxt.setColumns(25);
		
//		
	}
}
