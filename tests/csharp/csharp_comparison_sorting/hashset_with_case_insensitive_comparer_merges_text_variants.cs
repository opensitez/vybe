// vybe-test: csharp/csharp_comparison_sorting/hashset_with_case_insensitive_comparer_merges_text_variants
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

using System.Collections.Generic; var set = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase); set.Add("A"); set.Add("a"); __P((set.Count).ToString());
__Check("1");
