// vybe-test: csharp/csharp_dictionary_tryget_patterns/containskey_false_avoids_indexer_on_guard
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

using System.Collections.Generic; var map = new Dictionary<string, int>(); if (!map.ContainsKey("z")) __P(("skip").ToString());
__Check("skip");
