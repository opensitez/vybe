// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_reads_outer_instance_field
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Outer{int seed=4; public class Inner{Outer o; public Inner(Outer o){this.o=o;} public int Read()=>o.seed;} public int Via()=>new Inner(this).Read();} __P((new Outer().Via()).ToString());
__Check("4");
