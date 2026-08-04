// vybe-test: csharp/csharp_regex_basics/regex_options_ignore_case_matches_mixed
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

bool r=System.Text.RegularExpressions.Regex.IsMatch("Hello","hello",
    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
__P((r).ToString());
__Check("True");
