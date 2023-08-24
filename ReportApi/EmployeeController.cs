using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Cors;

namespace ReportApi
{
    public class Employee
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Salary { get; set; }
    }

    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class EmployeeController : ApiController
    {
        Employee[] employees = new Employee[]
        {
            new Employee{Id=0,Name="Morris",Salary="151110"},
            new Employee{Id=1,Name="John", Salary="120000"},
            new Employee{Id=2,Name="Chris",Salary="140000"},
            new Employee{Id=3,Name="Siraj", Salary="90000"}
        };

        public IEnumerable<Employee> Get()
        {
            return employees.ToList();
        }

        public Employee Get(int Id)
        {
            try
            {
                return employees[Id];
            }
            catch (Exception)
            {
                return new Employee();
            }
        }

    }
}