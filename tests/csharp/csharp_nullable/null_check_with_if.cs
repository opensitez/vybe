// vybe-test: csharp/csharp_nullable/null_check_with_if
// origin: languages/csharp/tests/csharp/test_csharp_nullable.rs

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

string s = null;
if (s == null) {
    __P(("is null").ToString());
} else {
    __P(("has value").ToString());
}
s = "test";
if (s != null) {
    __P(("has value").ToString());
}
__Check("is null\nhas value");
