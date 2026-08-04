// vybe-test: csharp/csharp_regex_advanced/replace_with_match_evaluator_transforms_each_match
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

string result = System.Text.RegularExpressions.Regex.Replace(
    "a1b2c3", @"\d",
    m => ((int.Parse(m.Value)*2)).ToString());
__P((result).ToString());
__Check("a2b4c6");
