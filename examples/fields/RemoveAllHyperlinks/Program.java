import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;

public class Program {
    public static void main(String[] args) throws Exception {
		//Open the Word template document.
		 WordDocument document = new WordDocument("Template.docx");
		// Find all hyperlink fields directly using findAllItemsByProperty
		ListSupport<Entity> hyperlinkFields = document.findAllItemsByProperty(EntityType.Field, "FieldType", FieldType.FieldHyperlink.toString());

		// Iterate backward through hyperlink fields
		for (int i = hyperlinkFields.getCount() - 1; i >= 0; i--) {
			WField wField = (WField) hyperlinkFields.get(i);
			removeHyperlinkFromField(wField);
		}

		hyperlinkFields.clear();
		//Saves the Word document.
		document.save("Output.docx", FormatType.Docx);
		document.close();
		System.out.println("Word document generated successfully");
    }
	static void removeHyperlinkFromField(WField field) {
        try {
            WParagraph paragraph = field.getOwnerParagraph();
            int itemIndex = paragraph.getChildEntities().indexOf(field);

            WTextRange textRange = new WTextRange(paragraph.getDocument());
            textRange.setText(getHyperlinkText(itemIndex, paragraph));

            paragraph.getChildEntities().removeAt(itemIndex);
            paragraph.getChildEntities().insert(itemIndex, textRange);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static String getHyperlinkText(int hyperlinkIndex, WParagraph paragraph) throws Exception {
        String text = "";
        java.util.Stack<Entity> fieldStack = new java.util.Stack<Entity>();
        fieldStack.push(paragraph.getChildEntities().get(hyperlinkIndex));

        boolean isFieldCode = true;
        int i = hyperlinkIndex + 1;

        while (i < paragraph.getItems().getCount()) {
            Entity item = paragraph.getChildEntities().get(i);

            if (item instanceof WField) {
                fieldStack.push(item);
                isFieldCode = true;
            } else if (item instanceof WFieldMark) {
                WFieldMark mark = (WFieldMark) item;
                if (mark.getType().getEnumValue() == FieldMarkType.FieldSeparator.getEnumValue()) {
                    isFieldCode = false;
                } else if (mark.getType().getEnumValue() == FieldMarkType.FieldEnd.getEnumValue()) {
                    if (fieldStack.size() == 1) {
                        fieldStack.clear();
                        return text;
                    } else {
                        fieldStack.pop();
                    }
                }
            } else if (!isFieldCode && item instanceof WTextRange) {
                text = text + ((WTextRange) item).getText();
            }

            i++;
        }

        return text;
    }
}
