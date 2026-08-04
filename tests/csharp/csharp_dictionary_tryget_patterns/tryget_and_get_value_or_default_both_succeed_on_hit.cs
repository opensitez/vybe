// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_and_get_value_or_default_both_succeed_on_hit
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

using System.Collections.Generic; var map = new Dictionary<string, int> { ["z"] = 44 }; bool ok = map.TryGetValue("z", out int t); int g = map.GetValueOrDefault("z"); __P((ok).ToString()); __P((g).ToString());
__Check("True\n44");
