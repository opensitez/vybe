// vybe-test: csharp/csharp_regex_balancing_group_definitions/regex_balancing_group_case_1

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

string pattern = @"^((?<open>\()|(?<-open>\))|[^()]+)+$(?(open)(?!))";
bool ok = System.Text.RegularExpressions.Regex.IsMatch("(test_1)", pattern);
__P(ok.ToString());
__Check("True");
