// vybe-test: csharp/csharp_nested_type_access/nested_access_sibling_nested_types_independent
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

class Duo{public class A{public int Bump(int n)=>n+1;} public class B{public int Bump(int n)=>n+2;}} __P((new Duo.A().Bump(5)).ToString()); __P((new Duo.B().Bump(5)).ToString());
__Check("6\n7");
