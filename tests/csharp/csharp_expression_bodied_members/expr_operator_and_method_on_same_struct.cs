// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_and_method_on_same_struct
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

struct Num { public int V; public static Num operator +(Num a, Num b) => new Num { V = a.V + b.V }; public int Double() => V * 2; }
var n = new Num { V = 3 } + new Num { V = 4 }; __P((n.Double()).ToString());
__Check("14");
