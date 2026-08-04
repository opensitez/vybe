// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

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

// for_loop_bounds
string feature = "for_loop_bounds"; __P((feature.Length > 0).ToString());
__Check("True");
