// vybe-test: csharp/csharp_map_set_collections/hashset_contains_reports_membership
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var set = new HashSet<string> { "alpha", "beta" }; __Check((set.Contains("beta")).ToString(), "True");
