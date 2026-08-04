// vybe-test: csharp/csharp_error_handling/finally_always_runs
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
    __P(("before").ToString());
    throw new Exception("err");
} catch (Exception e) {
    __P(("caught").ToString());
} finally {
    __P(("always").ToString());
}
__Check("before\ncaught\nalways");
