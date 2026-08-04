// vybe-test: csharp/csharp_pattern_deconstruct/var_pattern_binds_matched_value_regardless_of_type
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

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

object value = 42;
if (value is var captured) __P((captured).ToString());
__Check("42");
