// vybe-test: csharp/csharp_list_operations/add_range_appends_all_elements_of_collection
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>{1};
list.AddRange(new[]{2,3,4});
__Check((list.Count).ToString(), "4");
