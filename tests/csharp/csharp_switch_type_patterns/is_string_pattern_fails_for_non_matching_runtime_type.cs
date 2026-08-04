// vybe-test: csharp/csharp_switch_type_patterns/is_string_pattern_fails_for_non_matching_runtime_type
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
if (boxed is string text) {
    __P((text).ToString());
} else {
    __P(("not-string").ToString());
}
__Check("not-string");
