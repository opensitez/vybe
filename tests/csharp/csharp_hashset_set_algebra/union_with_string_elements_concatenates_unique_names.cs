// vybe-test: csharp/csharp_hashset_set_algebra/union_with_string_elements_concatenates_unique_names
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<string> { "a" }; a.UnionWith(new[] { "b", "a" }); __Check((a.Count).ToString(), "2");
