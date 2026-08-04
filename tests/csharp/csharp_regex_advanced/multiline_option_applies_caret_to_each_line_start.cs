// vybe-test: csharp/csharp_regex_advanced/multiline_option_applies_caret_to_each_line_start
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

var matches = System.Text.RegularExpressions.Regex.Matches(
    "start\nnew line", @"^[a-z]",
    System.Text.RegularExpressions.RegexOptions.Multiline);
__P((matches.Count).ToString());
__Check("2");
