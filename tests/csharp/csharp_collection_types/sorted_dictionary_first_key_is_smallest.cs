// vybe-test: csharp/csharp_collection_types/sorted_dictionary_first_key_is_smallest
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sd=new System.Collections.Generic.SortedDictionary<int,string>{{3,"c"},{1,"a"},{2,"b"}};
__Check((sd.Keys.First()).ToString(), "1");
