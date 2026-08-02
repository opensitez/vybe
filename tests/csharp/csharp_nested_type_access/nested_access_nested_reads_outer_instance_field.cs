// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_reads_outer_instance_field
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer{int seed=4; public class Inner{Outer o; public Inner(Outer o){this.o=o;} public int Read()=>o.seed;} public int Via()=>new Inner(this).Read();} __Check((new Outer().Via()).ToString(), "4");
