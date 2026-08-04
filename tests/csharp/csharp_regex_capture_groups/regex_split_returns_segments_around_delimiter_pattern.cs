// vybe-test: csharp/csharp_regex_capture_groups/regex_split_returns_segments_around_delimiter_pattern
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

var parts = System.Text.RegularExpressions.Regex.Split("one,two,three", ",");
__P((parts[1]).ToString());
__Check("two");
