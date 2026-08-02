// vybe-test: csharp/csharp_collection_types/sorted_set_maintains_unique_sorted_elements
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var s=new System.Collections.Generic.SortedSet<int>{3,1,4,1,5};
__Check((s.Count).ToString(), "4");
__Check((s.Min).ToString(), "1"); __Check((s.Max).ToString(), "5");
