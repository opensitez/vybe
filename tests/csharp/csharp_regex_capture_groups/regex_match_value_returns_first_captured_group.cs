// vybe-test: csharp/csharp_regex_capture_groups/regex_match_value_returns_first_captured_group
// origin: languages/csharp/tests/csharp/test_csharp_regex_capture_groups.rs

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

var match = System.Text.RegularExpressions.Regex.Match("id=42", @"id=(\d+)");
__P((match.Groups[1].Value).ToString());
__Check("42");
