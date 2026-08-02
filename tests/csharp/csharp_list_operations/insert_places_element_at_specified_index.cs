// vybe-test: csharp/csharp_list_operations/insert_places_element_at_specified_index
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>{1,3};
list.Insert(1, 2);
__Check((list[1]).ToString(), "2");
