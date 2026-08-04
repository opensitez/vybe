// vybe-test: csharp/exceptions_advanced/catch_when_filter
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

try {
    throw new Exception("error 42");
} catch (Exception e) when (e.Message.Contains("42")) {
    __P(("filtered catch: " + e.Message).ToString());
}
__Check("filtered catch: error 42");
