// vybe-test: csharp/csharp_string_culture/string_compare_invariant_culture_ignores_locale
// origin: languages/csharp/tests/csharp/test_csharp_string_culture.rs

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

int r=string.Compare("hello","HELLO",System.StringComparison.InvariantCultureIgnoreCase);
__P((r==0).ToString());
__Check("True");
