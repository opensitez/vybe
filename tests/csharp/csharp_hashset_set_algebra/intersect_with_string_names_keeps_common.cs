// vybe-test: csharp/csharp_hashset_set_algebra/intersect_with_string_names_keeps_common
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<string> { "x", "y" }; a.IntersectWith(new[] { "y", "z" }); __Check((a.Contains("y")).ToString(), "True"); __Check((a.Count).ToString(), "1");
