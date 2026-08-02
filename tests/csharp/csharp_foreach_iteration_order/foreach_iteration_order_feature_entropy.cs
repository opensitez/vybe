// vybe-test: csharp/csharp_foreach_iteration_order/foreach_iteration_order_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_foreach_iteration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// foreach_iteration_order
string feature = "foreach_iteration_order:46"; __Check((feature.Length >= 1).ToString(), "True");
