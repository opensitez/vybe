// vybe-test: csharp/modern_features/is_constant_pattern
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

object obj = null;
__P((obj is null).ToString());
obj = 42;
__P((obj is 42).ToString());
__P((obj is 43).ToString());
__Check("True\nTrue\nFalse");
