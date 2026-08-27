// vybe-test: csharp/csharp_regex_match_evaluator_lambdas/regex_match_evaluator_case_7

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

string text = "item 7 test";
string res = System.Text.RegularExpressions.Regex.Replace(text, @"\d+", m => (int.Parse(m.Value) * 2).ToString());
__P(res);
__Check("item 14 test");
