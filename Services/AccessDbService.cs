using System;
using System.Data.OleDb;
using Microsoft.Extensions.Configuration;
using TLM_Canteen.Models;

namespace TLM_Canteen.Services;

public class AccessDbService
{
    private readonly string _connectionString;

    public AccessDbService(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("AccessConnection");
    }

    public void GetDataCheckInOut()
    {
          
    }

}
