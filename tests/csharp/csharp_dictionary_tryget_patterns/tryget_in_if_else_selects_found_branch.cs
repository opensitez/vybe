// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_in_if_else_selects_found_branch
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

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

using System.Collections.Generic; var map = new Dictionary<string, int> { ["found"] = 11 }; if (map.TryGetValue("found", out int v)) __P(("yes:" + v).ToString()); else __P(("no").ToString());
__Check("yes:11");
