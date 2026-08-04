// vybe-test: csharp/csharp_dictionary_contracts/foreach_over_dictionary_emits_key_value_pairs_in_insertion_order
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_contracts.rs

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
var map = new Dictionary<string, int> {
    ["b"] = 2,
    ["a"] = 1,
    ["c"] = 3
};
foreach (var entry in map) {
    __P((entry.Key + ":" + entry.Value).ToString());
}
__Check("b:2\na:1\nc:3");
