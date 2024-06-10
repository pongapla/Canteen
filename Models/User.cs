namespace TLM_Canteen.Models;

public class SearchUser
{
    public int No { get; set; } 
    public string Code { get; set; }
    public string Name { get; set; }
    public string Department { get; set; }
    public string _DateTime { get; set; } 
    public string _Time { get; set; }
    public int? Total { get; set; }
}

public class CheckInout  
{
    public string Code { get; set; }
    public DateTime _DateTime { get; set; }
}