namespace TLM_Canteen.Models;

public class User
{
    public string UserName { get; set; }
    public string Department { get; set; }
    public DateTime? _Date { get; set; }
    public string Time {  get; set; }
    public int Total { get; set; }
}

public class CheckInout  
{
    public string Code { get; set; }
    public DateTime _DateTime { get; set; }
}