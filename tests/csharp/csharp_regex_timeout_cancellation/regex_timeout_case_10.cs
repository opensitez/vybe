// vybe-test: csharp/csharp_regex_timeout_cancellation/regex_timeout_case_10

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var regex = new System.Text.RegularExpressions.Regex("abc_10", System.Text.RegularExpressions.RegexOptions.None, TimeSpan.FromSeconds(2));
__P(regex.MatchTimeout.TotalSeconds.ToString());
__Check("2");
