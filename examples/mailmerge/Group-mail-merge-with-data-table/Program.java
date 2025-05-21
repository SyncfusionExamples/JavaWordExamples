import java.io.ByteArrayInputStream;
import com.syncfusion.docio.MailMergeDataTable;
import com.syncfusion.docio.MergeImageFieldEventArgs;
import com.syncfusion.docio.MergeImageFieldEventHandler;
import com.syncfusion.docio.WordDocument;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;
import com.syncfusion.javahelper.system.io.FileAccess;
import com.syncfusion.javahelper.system.io.FileMode;
import com.syncfusion.javahelper.system.io.FileStreamSupport;

public class Program {
    public static void main(String[] args) throws Exception 
	{
		//Loads an existing Word document into DocIO instance.
        WordDocument document = new WordDocument("Template.docx");
        //Gets the employee details as IEnumerable collection.
        ListSupport<Employee> employeeList = getEmployees();
        //Uses the mail merge events handler for image fields.
        document.getMailMerge().MergeImageField.add("mergeField_EmployeeImage", new MergeImageFieldEventHandler() {
        ListSupport<MergeImageFieldEventHandler> delegateList = new ListSupport<MergeImageFieldEventHandler>(
        MergeImageFieldEventHandler.class);
        //Represents event handling for MergeFieldEventHandlerCollection.
        public void invoke(Object sender, MergeImageFieldEventArgs args) throws Exception 
        {
        	mergeField_EmployeeImage(sender, args);
        }
        //Represents the method that handles MergeField event.
        public void dynamicInvoke(Object... args) throws Exception 
        {
        	mergeField_EmployeeImage((Object) args[0], (MergeImageFieldEventArgs) args[1]);
        }
        //Represents the method that handles MergeField event to add collection item.
        public void add(MergeImageFieldEventHandler delegate) throws Exception 
        {
        	if (delegate != null)
        		delegateList.add(delegate);
        }
        //Represents the method that handles MergeField event to remove collection item.
        public void remove(MergeImageFieldEventHandler delegate) throws Exception 
        {
        	if (delegate != null)
        		delegateList.remove(delegate);
        }
        });
        //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
        MailMergeDataTable dataSource = new MailMergeDataTable("Employees",employeeList);
        //Executes the mail merge for group.
        document.getMailMerge().executeGroup(dataSource);
        //Saves and closes the WordDocument instance.
        document.save("Sample.docx");
        document.close();
    }
    public static ListSupport<Employee> getEmployees()throws Exception
    {
    	ListSupport<Employee> employees = new ListSupport<Employee>(Employee.class);
    	employees.add(new Employee("Nancy","Smith","Sales Representative","505 - 20th Ave. E. Apt. 2A,","Seattle","WA","USA","Nancy.png"));
    	employees.add(new Employee("Andrew","Fuller","Vice President, Sales","908 W. Capital Way","Tacoma","WA","USA","Andrew.png"));
    	return employees;
    }

    public static void mergeField_EmployeeImage(Object sender, MergeImageFieldEventArgs args) throws Exception 
    {
    	//Binds image from file system during mail merge.
    	if ((args.getFieldName()).equals("Photo")) 
    	{
    		String ProductFileName = args.getFieldValue().toString();
    		//Gets the image from file system.
    		FileStreamSupport imageStream = new FileStreamSupport(ProductFileName, FileMode.Open, FileAccess.Read);
    		ByteArrayInputStream stream = new ByteArrayInputStream(imageStream.toArray());
    		args.setImageStream(stream);
    	}
    }	
}
