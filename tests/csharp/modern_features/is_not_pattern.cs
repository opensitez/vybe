// vybe-test: csharp/modern_features/is_not_pattern
// origin: languages/csharp/tests/csharp/test_modern_features.rs

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

object obj = "test";
if (obj is not null) {
    __P(("not null").ToString());
}
if (obj is not int) {
    __P(("not int").ToString());
}
__Check("not null\nnot int");
