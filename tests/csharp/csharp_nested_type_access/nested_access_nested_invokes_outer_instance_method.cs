// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_invokes_outer_instance_method
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

class Outer{int Double(int n)=>n*2; public class Inner{Outer o; public Inner(Outer o){this.o=o;} public int Run(int n)=>o.Double(n);} public int Via(int n)=>new Inner(this).Run(n);} __P((new Outer().Via(6)).ToString());
__Check("12");
