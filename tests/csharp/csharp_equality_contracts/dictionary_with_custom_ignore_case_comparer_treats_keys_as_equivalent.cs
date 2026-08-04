// vybe-test: csharp/csharp_equality_contracts/dictionary_with_custom_ignore_case_comparer_treats_keys_as_equivalent
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

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

using System.Collections.Generic;
var map = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
map["User"] = 1;
__P((map.ContainsKey("user")).ToString());
__P((map["USER"]).ToString());
__Check("True\n1");
