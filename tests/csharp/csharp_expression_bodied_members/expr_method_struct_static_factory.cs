// vybe-test: csharp/csharp_expression_bodied_members/expr_method_struct_static_factory
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

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

struct Point { public int X, Y; public static Point Origin() => new Point { X = 0, Y = 0 }; }
var p = Point.Origin();
__P((p.X).ToString()); __P((p.Y).ToString());
__Check("0\n0");
