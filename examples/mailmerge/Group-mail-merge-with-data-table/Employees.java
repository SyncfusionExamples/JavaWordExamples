public class Employee 
{
	private String _firstName;
	private String _lastName;
	private String _address;
	private String _city;
	private String _region;
	private String _country;
	private String _title;
	private String _photo;
	public String getFirstName()throws Exception
	{
		return _firstName;
	}
	public String setFirstName(String value)throws Exception
	{
		_firstName=value;
		return value;
	}
	public String getLastName()throws Exception
	{
		return _lastName;
	}
	public String setLastName(String value)throws Exception
	{
		_lastName=value;
		return value;
	}
	public String getAddress()throws Exception
	{
		return _address;
	}
	public String setAddress(String value)throws Exception
	{
		_address=value;
		return value;
	}
	public String getCity()throws Exception
	{
		return _city;
	}
	public String setCity(String value)throws Exception
	{
		_city=value;
		return value;
	}
	public String getRegion()throws Exception
	{
		return _region;
	}
	public String setRegion(String value)throws Exception
	{
		_region=value;
		return value;
	}
	public String getCountry()throws Exception{
		return _country;
	}
	public String setCountry(String value)throws Exception
	{
		_country=value;
		return value;
	}
	public String getTitle()throws Exception
	{
		return _title;
	}
	public String setTitle(String value)throws Exception
	{
		_title=value;
		return value;
	}
	public String getPhoto()throws Exception
	{
		return _photo;
	}
	public String setPhoto(String image)throws Exception
	{
		_photo=image;
		return image;
	}
	public Employee(String firstName,String lastName,String title,String address,String city,String region,String country,String photoFilePath)throws Exception
	{
		setFirstName(firstName);
		setLastName(lastName);
		setTitle(title);
		setAddress(address);
		setCity(city);
		setRegion(region);
		setCountry(country);
		setPhoto((photoFilePath));
	}
}