// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_struct_copy_independent
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

class Sheet{public struct Cell{public int V;} public int Sum(){var a=new Cell(); var b=a; a.V=3; b.V=5; return a.V+b.V;}} __P((new Sheet().Sum()).ToString());
__Check("8");
