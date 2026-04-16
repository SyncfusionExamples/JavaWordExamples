import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;

public class Program {
    public static void main(String[] args) throws Exception {
		//Open the Word template document.
		WordDocument document = new WordDocument("Template.docx");		
		//Find all fields in the Word document.
		ListSupport<Entity> fields = document.findAllItemsByProperty(EntityType.Field, null, null);
		//Remove hyperlink fields and replace with text.
		for(Object field : fields)
		{
			WField wField = (WField) field; 
			if(wField.getFieldType().getEnumValue()==FieldType.FieldHyperlink.getEnumValue())
			{
				RemoveHyperlink(wField);
			}
		}
		//Clear the temporary field collection.
		fields.clear();
		//Saves the Word document.
		document.save("Output.docx", FormatType.Docx);
		document.close();
		System.out.println("Word document generated successfully");
    }
	public static String GetHyperlinkText(int hyperlinkIndex, WParagraph paragraph) throws Exception {
    		
		String text = "";
		//Add the hyperlink field in stack to get the textrange from nested fields.
		Stack<Entity> fieldStack = new Stack<Entity>();
		fieldStack.push(paragraph.getChildEntities().get(hyperlinkIndex));
		//Flag to get the text from textrange between field separator and end.
		boolean isFieldCode = true;
		int i = (hyperlinkIndex + 1);
			while (i < paragraph.getItems().getCount())
			{
				Entity item = paragraph.getChildEntities().get(i);
				//If it is nested field, maintain in stack.
				if ((item instanceof WField))
				{
					fieldStack.push(item);
					//Set flag to skip getting text from textrange.
					isFieldCode = true;
				}
				else if ((item instanceof WFieldMark) && (((WFieldMark)(item)).getType().getEnumValue() == FieldMarkType.FieldSeparator.getEnumValue()))
					//If separator is reached, set flag to read text from textrange.
					isFieldCode = false;
				else if ((item instanceof WFieldMark) && (((WFieldMark)(item)).getType().getEnumValue() == FieldMarkType.FieldEnd.getEnumValue()))
				{
					//If field end is reached, check whether it is end of hyperlink field and skip the iteration.
					if (fieldStack.size() == 1)
					{
						fieldStack.clear();
						return text;
					}
					else
						fieldStack.pop();
				}
				else if (!isFieldCode && (item instanceof WTextRange))
					text = (text + ((WTextRange)(item)).getText());
				i = (i + 1);
			}
		return text;
	}

	//Method to replace a hyperlink field with a text value
	public static void RemoveHyperlink(WField field) throws Exception {
		WParagraph paragraph = field.getOwnerParagraph();
		int itemIndex = paragraph.getChildEntities().indexOf(field);
		WTextRange textRange = new WTextRange(paragraph.getDocument());
		//Gets the text from hyperlink field.
		textRange.setText(GetHyperlinkText(itemIndex, paragraph));
		//Removes the hyperlink field
		paragraph.getChildEntities().removeAt(itemIndex);
		//Inserts the hyperlink text
		paragraph.getChildEntities().insert(itemIndex, textRange);
	
	}
}
