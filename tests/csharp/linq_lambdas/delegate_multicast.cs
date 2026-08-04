// vybe-test: csharp/linq_lambdas/delegate_multicast
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

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

Action<string> logger = msg => __P(("LOG: " + msg).ToString());
Action<string> printer = msg => __P(("PRINT: " + msg).ToString());
Action<string> both = logger + printer;
both("hello");
__Check("LOG: hello\nPRINT: hello");
