// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_property_returns_field_by_readonly_ref
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

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

struct Point{public int X; public ref readonly int Rx=>ref X;} var p=new Point(); p.X=11; __P((p.Rx).ToString());
__Check("11");
