// vybe-test: csharp/exceptions_advanced/nested_try_catch
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
    __P(("outer try").ToString());
    try {
        throw new Exception("inner error");
    } catch (Exception e) {
        __P(("inner catch: " + e.Message).ToString());
    }
    __P(("after inner").ToString());
} catch (Exception) {
    __P(("outer catch").ToString());
}
__Check("outer try\ninner catch: inner error\nafter inner");
