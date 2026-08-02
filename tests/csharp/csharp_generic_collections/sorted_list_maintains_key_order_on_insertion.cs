// vybe-test: csharp/csharp_generic_collections/sorted_list_maintains_key_order_on_insertion
// origin: languages/csharp/tests/csharp/test_csharp_generic_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sl = new System.Collections.Generic.SortedList<int,string>();
sl.Add(3,"c"); sl.Add(1,"a"); sl.Add(2,"b");
__Check((sl.Keys[0]).ToString(), "1");
__Check((sl.Values[0]).ToString(), "a");
