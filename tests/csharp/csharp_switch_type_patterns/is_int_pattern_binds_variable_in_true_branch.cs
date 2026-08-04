// vybe-test: csharp/csharp_switch_type_patterns/is_int_pattern_binds_variable_in_true_branch
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

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

object boxed = 12;
if (boxed is int value) {
    __P((value + 1).ToString());
} else {
    __P(("no").ToString());
}
__Check("13");
