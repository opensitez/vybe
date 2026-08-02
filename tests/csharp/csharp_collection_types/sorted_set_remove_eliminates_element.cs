// vybe-test: csharp/csharp_collection_types/sorted_set_remove_eliminates_element
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var s=new System.Collections.Generic.SortedSet<int>{1,2,3};
s.Remove(2);
__Check((s.Count).ToString(), "2");
