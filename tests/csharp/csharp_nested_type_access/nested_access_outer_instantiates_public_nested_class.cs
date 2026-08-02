// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_instantiates_public_nested_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Shell{public class Core{public int Id=7;}} __Check((new Shell.Core().Id).ToString(), "7");
