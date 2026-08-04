// vybe-test: csharp/exceptions_advanced/catch_finally_on_error
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

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

string result = "start";
try {
    int x = 10 / 0;
    result = "never";
} catch (DivideByZeroException) {
    result = "caught";
} finally {
    result += " + finally";
}
__P((result).ToString());
__Check("caught + finally");
