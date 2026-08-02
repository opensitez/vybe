// vybe-test: csharp/csharp_list_operations/sort_orders_elements_ascending
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>{3,1,2};
list.Sort();
__Check((list[0]).ToString(), "1"); __Check((list[2]).ToString(), "3");
