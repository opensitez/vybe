// vybe-test: csharp/csharp_string_comparison/string_comparer_ordinal_ignore_case_used_in_sorted_set
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var set = new System.Collections.Generic.SortedSet<string>(
    System.StringComparer.OrdinalIgnoreCase);
set.Add("Apple"); set.Add("apple");
__Check((set.Count).ToString(), "1");
