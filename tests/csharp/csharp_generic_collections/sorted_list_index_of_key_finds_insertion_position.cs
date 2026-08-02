// vybe-test: csharp/csharp_generic_collections/sorted_list_index_of_key_finds_insertion_position
// origin: languages/csharp/tests/csharp/test_csharp_generic_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sl = new System.Collections.Generic.SortedList<string,int>{{"a",1},{"b",2},{"c",3}};
__Check((sl.IndexOfKey("b")).ToString(), "1");
