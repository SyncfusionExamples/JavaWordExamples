import java.io.File;
import java.nio.file.*;
import java.util.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.javahelper.system.data.*;

public class ReplaceMergeFieldWithHTML {
	private static final String DATA_DIR="resources\\replacemergefieldwithhtml";
	static HashMap<WParagraph, HashMap<Integer, String>> paraToInsertHTML = new HashMap<>();

	public static void main(String[] args) throws Exception {
		// Opens the template document.
        WordDocument document = new WordDocument(getDataPath("Template.docx"));
        
        // Creates mail merge events handler to replace merge field with HTML.
        document.getMailMerge().MergeField.add("mergeFieldEvent", new MergeFieldEventHandler() {
        	ListSupport<MergeFieldEventHandler> delegateList = new ListSupport<MergeFieldEventHandler>(
        			MergeImageFieldEventHandler.class);
        	@Override
            public void invoke(Object sender, MergeFieldEventArgs args) throws Exception{
                mergeFieldEvent(sender, args);
            }

			@Override
			public void add(MergeFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.add(delegate);
			}

			@Override
			public void dynamicInvoke(Object... args) throws Exception {
				 mergeFieldEvent((Object) args[0], (MergeFieldEventArgs) args[1]);
			}

			@Override
			public void remove(MergeFieldEventHandler delegate) throws Exception {
				if (delegate != null)
					delegateList.remove(delegate);
			}
        });

        // Gets data to perform mail merge.
        DataTableSupport table = getDataTable();
        // Performs the mail merge.
        document.getMailMerge().execute(table);
        // Append HTML to paragraph.
        insertHtml();
        // Saves the Word document instance.
        document.save("Output.docx");
        //Closes the Word document.
        document.close();

	}
	// Helper methods
    public static void mergeFieldEvent(Object sender, MergeFieldEventArgs args) throws Exception {
        if (args.getTableName().equals("HTML")) {
            if (args.getFieldName().equals("ProductList")) {
                // Gets the current merge field owner paragraph.
                WParagraph paragraph = args.getCurrentMergeField().getOwnerParagraph();
                // Gets the current merge field index in the current paragraph.
                int mergeFieldIndex = paragraph.getChildEntities().indexOf(args.getCurrentMergeField());
                // Maintain HTML in collection.
                HashMap<Integer, String> fieldValues = new HashMap<>();
                fieldValues.put(mergeFieldIndex, args.getFieldValue().toString());
                // Maintain paragraph in collection.
                paraToInsertHTML.put(paragraph, fieldValues);
                // Set field value as empty.
                args.setText("");
            }
        }
    }

    private static DataTableSupport getDataTable() throws Exception{
        DataTableSupport dataTable = new DataTableSupport("HTML");
        dataTable.getColumns().add("CustomerName");
        dataTable.getColumns().add("Address");
        dataTable.getColumns().add("Phone");
        dataTable.getColumns().add("ProductList");
        DataRowSupport datarow = dataTable.newRow();
        dataTable.getRows().add(datarow);
        datarow.set("CustomerName", "Nancy Davolio");
        datarow.set("Address", "59 rue de I'Abbaye, Reims 51100, France");
        datarow.set("Phone", "1-888-936-8638");
        
        // Reads HTML string from the file.
        String htmlString =new String( Files.readAllBytes(Paths.get(getDataPath("File.html")))); 
        // Read the HTML file and assign it to htmlString
        
        datarow.set("ProductList", htmlString);
        return dataTable;
    }

    private static void insertHtml() throws Exception{
        // Iterates through each item in the map.
        for (Map.Entry<WParagraph, HashMap<Integer, String>> entry : paraToInsertHTML.entrySet()) {
            WParagraph paragraph = entry.getKey();
            HashMap<Integer, String> values = entry.getValue();
            // Iterates through each value in the map.
            for (Map.Entry<Integer, String> valuePair : values.entrySet()) {
                int index = valuePair.getKey();
                String fieldValue = valuePair.getValue();
                // Inserts HTML string at the same position of mergefield in Word document.
                paragraph.getOwnerTextBody().insertXHTML(fieldValue, paragraph.getOwnerTextBody().getChildEntities().indexOf(paragraph), index);
            }
        }
        paraToInsertHTML.clear();
    }
    /**
	 * Get the file path
	 * 
	 * @param path specifies the file path
	 */
	public static String getDataPath(String path) {
		//Get the current directory.
        File dir = new File(System.getProperty("user.dir"));
        // Get the input folder.
        dir = new File(dir, DATA_DIR);
        //Get the file path
        dir = new File(dir, path);
        return dir.toString();
    }
}
