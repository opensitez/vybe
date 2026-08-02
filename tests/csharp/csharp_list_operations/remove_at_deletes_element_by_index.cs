// vybe-test: csharp/csharp_list_operations/remove_at_deletes_element_by_index
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<string>{"a","b","c"};
list.RemoveAt(0);
__Check((list[0]).ToString(), "b");
