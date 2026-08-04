// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_bitwise_or_on_flags
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

struct Bits { public int V; public static Bits operator |(Bits a, Bits b) => new Bits { V = a.V | b.V }; }
__P(((new Bits { V = 1 } | new Bits { V = 2 }).V).ToString());
__Check("3");
