// vybe-test: csharp/basics/boolean_and_or
// origin: languages/csharp/tests/csharp/test_basics.rs

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

__P((true && false).ToString());
        __P((true || false).ToString());
        __P((!true).ToString());
__Check("False\nTrue\nFalse");
