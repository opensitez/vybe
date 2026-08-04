// vybe-test: csharp/csharp_expression_bodied_members/expr_method_long_arithmetic
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

class Wide { public long Add(long a, long b) => a + b; }
__P((new Wide().Add(9000000000L, 1L)).ToString());
__Check("9000000001");
