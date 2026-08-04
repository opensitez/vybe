// vybe-test: csharp/strings_advanced/string_indexof_lastindexof
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

string s = "abcabc";
__P((s.IndexOf("bc")).ToString());
__P((s.LastIndexOf("bc")).ToString());
__P((s.IndexOf("xyz")).ToString());
__Check("1\n4\n-1");
