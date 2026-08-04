// vybe-test: csharp/strings_advanced/string_trim
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

string s = "  hello  ";
__P(("'" + s.Trim() + "'").ToString());
__P(("'" + s.TrimStart() + "'").ToString());
__P(("'" + s.TrimEnd() + "'").ToString());
__Check("'hello'\n'hello  '\n'  hello'");
