// vybe-test: csharp/csharp_sorted_collections/sorted_set_contains_finds_member
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ss = new SortedSet<string> { "a", "b" }; __Check((ss.Contains("b")).ToString(), "True");
