// vybe-test: csharp/strings_advanced/string_compareto
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

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

string a = "apple";
string b = "banana";
__P((a.CompareTo(b) < 0).ToString());
__P((b.CompareTo(a) > 0).ToString());
__P((a.CompareTo(a) == 0).ToString());
__Check("True\nTrue\nTrue");
