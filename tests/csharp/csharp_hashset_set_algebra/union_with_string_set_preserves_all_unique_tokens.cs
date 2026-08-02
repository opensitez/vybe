// vybe-test: csharp/csharp_hashset_set_algebra/union_with_string_set_preserves_all_unique_tokens
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<string> { "one", "two" }; a.UnionWith(new[] { "two", "three" }); __Check((a.Count).ToString(), "3");
