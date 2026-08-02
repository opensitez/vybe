// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_exposes_nested_via_property
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Shell{public class Core{public int Id=2;} Core _c=new Core(); public Core Inner=>_c;} __Check((new Shell().Inner.Id).ToString(), "2");
