// vybe-test: csharp/csharp_error_handling/try_finally
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

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

try {
    __P(("try").ToString());
} finally {
    __P(("finally").ToString());
}
__Check("try\nfinally");
