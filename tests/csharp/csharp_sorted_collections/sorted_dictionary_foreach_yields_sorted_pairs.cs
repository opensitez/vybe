// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_foreach_yields_sorted_pairs
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

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

using System.Collections.Generic; var sd = new SortedDictionary<string, int> { ["b"] = 2, ["a"] = 1, ["c"] = 3 }; foreach (var p in sd) __P((p.Key + ":" + p.Value).ToString());
__Check("a:1\nb:2\nc:3");
