// vybe-test: csharp/csharp_regex_advanced/matches_returns_all_non_overlapping_occurrences
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

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

var matches = System.Text.RegularExpressions.Regex.Matches("a1 b2 c3", @"\d");
__P((matches.Count).ToString());
__Check("3");
