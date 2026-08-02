// vybe-test: csharp/csharp_foreach_iteration_order/foreach_iteration_order_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_foreach_iteration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// foreach_iteration_order
double seed = 46; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
