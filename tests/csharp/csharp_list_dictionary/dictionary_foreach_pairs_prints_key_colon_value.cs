// vybe-test: csharp/csharp_list_dictionary/dictionary_foreach_pairs_prints_key_colon_value
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

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

using System.Collections.Generic; var map = new Dictionary<string, int> { ["b"] = 2, ["a"] = 1 }; foreach (var pair in map) __P((pair.Key + ":" + pair.Value).ToString());
__Check("b:2\na:1");
