// vybe-test: csharp/csharp_comparison_sorting/hashset_with_case_insensitive_comparer_merges_text_variants
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var set = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase); set.Add("A"); set.Add("a"); __Check((set.Count).ToString(), "1");
