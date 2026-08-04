// vybe-test: csharp/csharp_regex_basics/regex_match_extracts_first_occurrence
// origin: languages/csharp/tests/csharp/test_csharp_regex_basics.rs

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

var m=System.Text.RegularExpressions.Regex.Match("abc123def","[0-9]+");
__P((m.Value).ToString());
__Check("123");
