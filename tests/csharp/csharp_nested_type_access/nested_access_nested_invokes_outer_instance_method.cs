// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_invokes_outer_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer{int Double(int n)=>n*2; public class Inner{Outer o; public Inner(Outer o){this.o=o;} public int Run(int n)=>o.Double(n);} public int Via(int n)=>new Inner(this).Run(n);} __Check((new Outer().Via(6)).ToString(), "12");
