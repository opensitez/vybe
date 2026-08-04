// vybe-test: csharp/csharp_modern/boolean_values
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

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

bool t = true;
bool f = false;
__P((t).ToString());
__P((f).ToString());
__P((t && f).ToString());
__P((t || f).ToString());
__P((!t).ToString());
__Check("True\nFalse\nFalse\nTrue\nFalse");
