// vybe-test: csharp/csharp_list_operations/remove_all_deletes_every_matching_element
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>{1,2,3,4,5};
list.RemoveAll(x => x % 2 == 0);
__Check((list.Count).ToString(), "3");
