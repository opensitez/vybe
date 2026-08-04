// vybe-test: csharp/csharp_switch_expression_core/switch_expr_when_guard_false_skips_to_next_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

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

var x=3; __P((x switch{int n when n>10=>"big",int n when n>1=>"mid",_=>"small"}).ToString());
__Check("mid");
