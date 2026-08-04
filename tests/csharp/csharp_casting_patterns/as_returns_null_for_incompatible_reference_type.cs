// vybe-test: csharp/csharp_casting_patterns/as_returns_null_for_incompatible_reference_type
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

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

object o=42;
string s=o as string;
__P((s==null).ToString());
__Check("True");
