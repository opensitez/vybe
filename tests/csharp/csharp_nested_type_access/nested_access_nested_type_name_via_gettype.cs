// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_type_name_via_gettype
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer{public class Inner{}} __Check((typeof(Outer.Inner).Name).ToString(), "Inner");
