import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;

public class Program {
    public static void main(String[] args) throws Exception 
	{
		//Loads an existing Word document into DocIO instance.
        WordDocument document = new WordDocument("Template.docx");
        //Get all fields present in a document.
		ListSupport<Entity> fields = document.findAllItemsByProperty(EntityType.Field, null, null);
		//Unlink all the fields.
		for (int i = 0; i < fields.getCount(); i++) {
			WField field = (WField) fields.get(i);
			if (field.getOwner() != null)
				  field.unlink();
		}
        //Saves and closes the WordDocument instance.
        document.save("Sample.docx");
        document.close();
    }
}
