// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_struct_unary_minus
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

struct Signed { public int V; public static Signed operator -(Signed s) => new Signed { V = -s.V }; }
__P(((-new Signed { V = 7 }).V).ToString());
__Check("-7");
