// vybe-test: csharp/csharp_list_operations/index_of_returns_first_position_of_element
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>{5,10,5};
__Check((list.IndexOf(5)).ToString(), "0");
