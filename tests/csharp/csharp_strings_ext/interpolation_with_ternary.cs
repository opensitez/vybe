// vybe-test: csharp/csharp_strings_ext/interpolation_with_ternary
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

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

int x = 5;
__P(($"x is {(x > 3 ? "big" : "small")}").ToString());
__Check("x is big");
