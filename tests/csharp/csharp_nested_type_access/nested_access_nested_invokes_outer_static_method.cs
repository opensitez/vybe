// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_invokes_outer_static_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer{static int Triple(int n)=>n*3; public class Inner{public int Run(int n)=>Triple(n);}} __Check((new Outer.Inner().Run(2)).ToString(), "6");
