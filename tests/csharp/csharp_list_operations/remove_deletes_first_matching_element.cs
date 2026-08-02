// vybe-test: csharp/csharp_list_operations/remove_deletes_first_matching_element
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>{1,2,2,3};
list.Remove(2);
__Check((list.Count).ToString(), "3"); __Check((list[1]).ToString(), "2");
