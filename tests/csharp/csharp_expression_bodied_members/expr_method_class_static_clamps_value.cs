// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_static_clamps_value
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

static class ClampUtil { public static int Clamp(int v, int lo, int hi) => v < lo ? lo : v > hi ? hi : v; }
__P((ClampUtil.Clamp(15, 0, 10)).ToString());
__Check("10");
