// vybe-test: csharp/csharp_nested_type_access/nested_access_fully_qualified_name_from_outside
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A{public class B{public int N=11;}} __Check((new A.B().N).ToString(), "11");
