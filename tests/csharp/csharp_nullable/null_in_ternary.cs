// vybe-test: csharp/csharp_nullable/null_in_ternary
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
__P((s != null ? s : "none").ToString());
s = "found";
__P((s != null ? s : "none").ToString());
__Check("none\nfound");
