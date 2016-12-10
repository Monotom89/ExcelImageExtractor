
import javafx.application.*;
import static javafx.application.Application.launch;
import javafx.stage.*;
import javafx.scene.*;
import javafx.scene.layout.*;
import javafx.scene.control.*;
import javafx.geometry.*;
import javafx.event.*;
import java.time.*;
import javafx.scene.paint.*;
import java.io.*;
import java.nio.file.*;
import java.io.*;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipEntry;





public class ExcelImageExtractor extends Application
{
    public static void main(String[] args)
    {
    	launch(args);
    }

    @Override
    public void start(Stage hauptFenster)
    {
        GridPane gp = new GridPane();
        // gp.setGridLinesVisible(true); //linien anzeigen, zur verdeutlichung
        gp.setPadding(new Insets(20.0));
        gp.setAlignment(Pos.CENTER);

        Label lbTop = new Label();
        lbTop.setText("Noch keine Excel-Datei ausgewählt");
        GridPane.setRowIndex(lbTop, 0);
        GridPane.setColumnIndex(lbTop,  0);
        GridPane.setMargin(lbTop, new Insets(10.0));
        GridPane.setColumnSpan(lbTop,  2);
        gp.getChildren().add(lbTop);

    	Button buLeft = new Button("Excel Datei auswählen");
        buLeft.setOnAction( ( (ev) -> auswahlDatei(hauptFenster, lbTop)) );

        GridPane.setRowIndex(buLeft, 1);
        GridPane.setColumnIndex(buLeft,  0);
        GridPane.setMargin(buLeft, new Insets(10.0));
        gp.getChildren().add(buLeft);

        Button buRight = new Button("Bilder extrahieren");
		buRight.setOnAction( ( (ev) -> bilderExtrahieren(hauptFenster, lbTop)) );
		GridPane.setRowIndex(buRight, 1);
        GridPane.setColumnIndex(buRight,  1);
        GridPane.setMargin(buRight, new Insets(10.0));
        gp.getChildren().add(buRight);

		hauptFenster.setTitle("Menu");
		hauptFenster.setScene(new Scene(gp, 500, 100));
		hauptFenster.show();
    }

    private void auswahlDatei(Stage hf, Label lb)
    {
        FileChooser fc = new FileChooser();
        FileChooser.ExtensionFilter ef = new FileChooser.ExtensionFilter("Excel-Dateien (*.xlsx)", "*.xlsx");
        fc.getExtensionFilters().add(ef);

        File datei = fc.showOpenDialog(hf);
        String dateiPfad = datei.toString();

        lb.setText(dateiPfad);
    }
    
    private void bilderExtrahieren(Stage hf, Label lb)
    {
    	//Neuen Unterordner erstellen aus Dateinamen
    	String ExcelDatei = lb.getText();
    	String DateiName = ExcelDatei.substring(ExcelDatei.lastIndexOf(File.separator)+1, ExcelDatei.indexOf("."));
    	String parentDir = ExcelDatei.substring(0, ExcelDatei.lastIndexOf(File.separator));
    	new File(parentDir + "/" + DateiName + " Extracted Images").mkdir();
    	
    	//File duplizieren, damit die alte Excel bestehen bleibt
    	try {
			DateiKopieren(new File(ExcelDatei), new File(parentDir + "/" + DateiName + " Extracted Images/ExcelKopie.xlsx"));
		} catch (IOException e)
    	{
			System.out.printf("%s%n",  e);
		}
    			
    	//Excel umbennenen
    	File ExcelAlt = new File(parentDir + "/" + DateiName + " Extracted Images/ExcelKopie.xlsx");
    	File ExcelNeu = new File(ExcelAlt.toString().substring(0, ExcelAlt.toString().indexOf(".")) + ".zip");
    	boolean success = ExcelAlt.renameTo(ExcelNeu);
    	if(success)
    	{
    		System.out.println("Excel umbenannt");
    	}

    	//unzippen
    	try
    	{
    		unZip(ExcelNeu.toString(), parentDir + "/" + DateiName + " Extracted Images/");
    		System.out.println("Excel unzipt");

    	}
    	catch(Exception e)
    	{
    		System.out.printf("%s%n",  e);
    	}
    	
    	//Bilder kopieren
    	File BilderOrdner = new File(parentDir + "/" + DateiName + " Extracted Images/xl/media");
    	File ParentDirFile = new File(parentDir + "/");
    	File[] Bilder = BilderOrdner.listFiles();
    	if(Bilder != null) //nur in die Schleife gehen, wenn der Bilderordner nicht leer ist
    		for (File child : Bilder)
    		{
    			try
    			{
					DateiKopieren(child, new File(ParentDirFile + child.toString().substring(child.toString().lastIndexOf(File.separator), child.toString().length())));
		    		System.out.println("Bilder kopiert");

    			} catch (IOException e)
    			{
		    		System.out.printf("%s%n",  e);
				}
    		}
    	
    	//Arbeitsordner wieder löschen
    	File Arbeitsordner = new File(parentDir + "/" + DateiName + " Extracted Images");
		
		 try
		 {
            delete(Arbeitsordner);
     		System.out.println("Arbeitsordner entfernt");

         }
		 catch(IOException e)
		 {
             e.printStackTrace();
             System.exit(0);
         }
    	
    	
    		
    	
    }
      
    private static void DateiKopieren(File source, File dest) throws IOException {
    //Übergeben werden muss der komplette neue Filename, nicht nur das Parent-Directory!	
    	
        InputStream is = null;
        OutputStream os = null;
        try {
            is = new FileInputStream(source);
            os = new FileOutputStream(dest);
            byte[] buffer = new byte[1024];
            int length;
            while ((length = is.read(buffer)) > 0) {
                os.write(buffer, 0, length);
            }
        } finally {
            is.close();
            os.close();
        }
    }
    
    public void unZip(String zipFile, String outputFolder){

        byte[] buffer = new byte[1024];

        try
        {

	       	//create output directory is not exists
	       	File folder = new File(outputFolder);
	       	if(!folder.exists())
	       		folder.mkdir();
	       	
	       	//get the zip file content
	       	ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFile));
	       	//get the zipped file list entry
	       	ZipEntry ze = zis.getNextEntry();
	       	while(ze!=null)
	       	{
	       	   String fileName = ze.getName();
	              File newFile = new File(outputFolder + File.separator + fileName);
	               //create all non exists folders
	               //else you will hit FileNotFoundException for compressed folder
	               new File(newFile.getParent()).mkdirs();
	               FileOutputStream fos = new FileOutputStream(newFile);
	               int len;
	               while ((len = zis.read(buffer)) > 0) 
	            	   fos.write(buffer, 0, len);
	               fos.close();
	               ze = zis.getNextEntry();
	       	}
		       zis.closeEntry();
		       zis.close();
        }
       catch(IOException ex)
       {
          ex.printStackTrace();
       }
      }
   
    public static void delete(File file) throws IOException
    {
		if(file.isDirectory())//directory is empty, then delete it
		{
			if(file.list().length==0)
			{
			   file.delete();
			}
			else
			{
			   //list all the directory contents
	    	   String files[] = file.list();
	    	   for (String temp : files)
	    	   {
	    	      //construct the file structure
	    	      File fileDelete = new File(file, temp);
	    	      //recursive delete
	    	     delete(fileDelete);
	    	   }
	    	   //check the directory again, if empty then delete it
	    	   if(file.list().length==0)
	    	   {
	       	     file.delete();
	    	   }
			}
		}
		else
		{
			file.delete();
		}
    }

}
