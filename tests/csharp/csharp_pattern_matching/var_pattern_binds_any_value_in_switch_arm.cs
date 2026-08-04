// vybe-test: csharp/csharp_pattern_matching/var_pattern_binds_any_value_in_switch_arm
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

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

object o = 42;
string result = o switch { var x when x is int n && n > 10 => "big int", _ => "other" };
__P((result).ToString());
__Check("big int");
