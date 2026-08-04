// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_implicit_conversion_to_int
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

struct Wrap { public int V; public static implicit operator int(Wrap w) => w.V; }
Wrap w = new Wrap { V = 42 }; int n = w; __P((n).ToString());
__Check("42");
