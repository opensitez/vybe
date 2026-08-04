// vybe-test: csharp/csharp_comparison_sorting/dictionary_with_case_insensitive_comparer_finds_key_variant
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

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

using System.Collections.Generic; var map = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase) { ["Key"] = 3 }; __P((map.ContainsKey("key")).ToString());
__Check("True");
