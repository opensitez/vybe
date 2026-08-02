// vybe-test: csharp/csharp_bcl_collections/hashset_add_returns_false_for_duplicate_element
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var set = new System.Collections.Generic.HashSet<int>();
__Check((set.Add(1)).ToString(), "True");
__Check((set.Add(1)).ToString(), "False");
