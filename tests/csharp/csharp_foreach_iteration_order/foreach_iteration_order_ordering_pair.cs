// vybe-test: csharp/csharp_foreach_iteration_order/foreach_iteration_order_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_foreach_iteration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// foreach_iteration_order
int seed = 46; int right = seed + 1; __Check((seed < right).ToString(), "True");
