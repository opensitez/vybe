// vybe-test: csharp/csharp_regex_non_backtracking_engine/regex_non_backtracking_case_10

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

var opt = System.Text.RegularExpressions.RegexOptions.NonBacktracking;
var regex = new System.Text.RegularExpressions.Regex("item_10", opt);
__P(regex.IsMatch("prefix_item_10_suffix").ToString());
__Check("True");
