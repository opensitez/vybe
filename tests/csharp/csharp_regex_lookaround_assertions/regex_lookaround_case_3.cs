// vybe-test: csharp/csharp_regex_lookaround_assertions/regex_lookaround_case_3

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

string input = "USD300";
var m = System.Text.RegularExpressions.Regex.Match(input, @"(?<=USD)\d+");
__P(m.Value);
__Check("300");
