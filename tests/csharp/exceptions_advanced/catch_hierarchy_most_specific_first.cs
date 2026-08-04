// vybe-test: csharp/exceptions_advanced/catch_hierarchy_most_specific_first
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
    throw new InvalidOperationException("invalid op");
} catch (InvalidOperationException e) {
    __P(("specific: " + e.Message).ToString());
} catch (Exception) {
    __P(("generic").ToString());
}
__Check("specific: invalid op");
