// vybe-test: csharp/csharp_operators/logical_operators
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

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

__P((true && true).ToString());
__P((true && false).ToString());
__P((false || true).ToString());
__P((false || false).ToString());
__P((!true).ToString());
__Check("True\nFalse\nTrue\nFalse\nFalse");
