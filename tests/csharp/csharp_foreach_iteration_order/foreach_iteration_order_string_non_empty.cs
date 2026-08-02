// vybe-test: csharp/csharp_foreach_iteration_order/foreach_iteration_order_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_foreach_iteration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// foreach_iteration_order
string feature = "foreach_iteration_order"; __Check((feature.Length > 0).ToString(), "True");
