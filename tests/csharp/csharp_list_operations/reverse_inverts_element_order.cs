// vybe-test: csharp/csharp_list_operations/reverse_inverts_element_order
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>{1,2,3};
list.Reverse();
__Check((list[0]).ToString(), "3");
