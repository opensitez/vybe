// vybe-test: csharp/csharp_null_propagation/null_coalescing_operator_selects_right_when_left_is_null
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

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

string value = null; __P((value ?? "right").ToString());
__Check("right");
