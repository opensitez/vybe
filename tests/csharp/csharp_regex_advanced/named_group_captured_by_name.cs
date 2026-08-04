// vybe-test: csharp/csharp_regex_advanced/named_group_captured_by_name
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

var m = System.Text.RegularExpressions.Regex.Match("date=2024-06-15", @"(?<year>\d{4})-(?<month>\d{2})");
__P((m.Groups["year"].Value).ToString());
__P((m.Groups["month"].Value).ToString());
__Check("2024\n06");
