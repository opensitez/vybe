// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_captures_outer_readonly
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer{public readonly int Seed=10; public class Inner{Outer o; public Inner(Outer o){this.o=o;} public int Read()=>o.Seed;} public int Via()=>new Inner(this).Read();} __Check((new Outer().Via()).ToString(), "10");
