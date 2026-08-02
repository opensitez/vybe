// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_reads_outer_static_field
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer{static int tally=8; public class Inner{public int Read()=>tally;}} __Check((new Outer.Inner().Read()).ToString(), "8");
