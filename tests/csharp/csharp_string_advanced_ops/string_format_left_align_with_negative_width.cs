// vybe-test: csharp/csharp_string_advanced_ops/string_format_left_align_with_negative_width
// origin: languages/csharp/tests/csharp/test_csharp_string_advanced_ops.rs

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

__P((string.Format("{0,-10}|","hello")).ToString());
__Check("hello     |");
