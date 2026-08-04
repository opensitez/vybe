// vybe-test: csharp/exceptions_advanced/multiple_catch_blocks
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
    throw new ArgumentException("bad arg");
} catch (ArgumentNullException) {
    __P(("null").ToString());
} catch (ArgumentException e) {
    __P(("arg: " + e.Message).ToString());
} catch (Exception) {
    __P(("generic").ToString());
}
__Check("arg: bad arg");
