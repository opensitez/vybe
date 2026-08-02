// vybe-test: csharp/csharp_foreach_iteration_order/foreach_iteration_order_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_foreach_iteration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// foreach_iteration_order
var values = new System.Collections.Generic.List<int> { 46, 47, 46 }; __Check((values.Count == 3).ToString(), "True");
