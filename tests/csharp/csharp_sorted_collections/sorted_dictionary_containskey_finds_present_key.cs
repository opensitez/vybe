// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_containskey_finds_present_key
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [2] = "two" }; __Check((sd.ContainsKey(2)).ToString(), "True");
