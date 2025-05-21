import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.collections.generic.DictionarySupport;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

public class Program {
    public static void main(String[] args) throws Exception {
        // JSON string containing nested organizational data
    	String json = "{\"Organizations\":[{\"BranchName\":\"UK Office\",\"Address\":\"120 Hanover Sq.\",\"City\":\"London\",\"ZipCode\":\"WX1 6LT\",\"Country\":\"UK\",\"Departments\":[{\"DepartmentName\":\"Marketing\",\"Supervisor\":\"Nancy Davolio\",\"Employees\":[{\"EmployeeName\":\"Thomas Hardy\",\"EmployeeID\":\"1001\",\"JoinedDate\":\"05/27/1996\"},{\"EmployeeName\":\"Maria Anders\",\"EmployeeID\":\"1002\",\"JoinedDate\":\"04/10/1998\"}]},{\"DepartmentName\":\"Production\",\"Supervisor\":\"Andrew Fuller\",\"Employees\":[{\"EmployeeName\":\"Elizabeth Lincoln\",\"EmployeeID\":\"1003\",\"JoinedDate\":\"05/15/1996\"},{\"EmployeeName\":\"Antonio Moreno\",\"EmployeeID\":\"1004\",\"JoinedDate\":\"04/22/1996\"}]}]}]}";

        // Convert JSON string into a JsonObject
        JsonObject data = new Gson().fromJson(json, JsonObject.class);

        // Convert JsonObject to DictionarySupport (required format for mail merge)
        DictionarySupport<String, Object> result = getData(data);

        // Load the Word document template
        WordDocument document = new WordDocument("Template.docx");

        // Extract the list of organizations for mail merge
        MailMergeDataTable dataTable = new MailMergeDataTable("Organizations", 
            (ListSupport<Object>) result.get("Organizations"));

        // Perform nested mail merge with the data table
        document.getMailMerge().executeNestedGroup(dataTable);

        // Save the generated Word document
        document.save("Result.docx", FormatType.Docx);
        document.close();

        System.out.println("Word document generated successfully.");
    }

    // Converts a JsonObject into a DictionarySupport for Syncfusion mail merge
    public static DictionarySupport<String, Object> getData(JsonObject data) throws Exception {
        DictionarySupport<String, Object> map = new DictionarySupport<>(String.class, Object.class);

        for (String key : data.keySet()) {
            Object keyValue = null;
            if (data.get(key).isJsonArray()) {
                // Convert array to ListSupport
                keyValue = getData(data.getAsJsonArray(key));
            } else if (data.get(key).isJsonPrimitive()) {
                // Use primitive value directly
                keyValue = data.get(key).getAsString();
            }
            map.add(key, keyValue);
        }
        return map;
    }

    // Converts a JsonArray into ListSupport recursively
    public static ListSupport<Object> getData(JsonArray arr) throws Exception {
        ListSupport<Object> list = new ListSupport<>(Object.class);

        for (int i = 0; i < arr.size(); i++) {
            Object keyValue = null;
            if (arr.get(i).isJsonObject()) {
                // Recursively convert nested objects
                keyValue = getData(arr.get(i).getAsJsonObject());
            }
            list.add(keyValue);
        }
        return list;
    }
}
