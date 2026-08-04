// vybe-test: csharp/modern_features/null_conditional_chain
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

class Inner { public string Value = "found"; }
class Outer { public Inner Child; }
var o = new Outer();
__P((o.Child?.Value ?? "missing").ToString());
o.Child = new Inner();
__P((o.Child?.Value ?? "missing").ToString());
__Check("missing\nfound");
