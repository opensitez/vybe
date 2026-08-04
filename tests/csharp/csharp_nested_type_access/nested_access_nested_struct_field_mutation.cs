// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_struct_field_mutation
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

class Canvas{public struct Dot{public int X;} public Dot Make(){var d=new Dot(); d.X=9; return d;}} __P((new Canvas().Make().X).ToString());
__Check("9");
