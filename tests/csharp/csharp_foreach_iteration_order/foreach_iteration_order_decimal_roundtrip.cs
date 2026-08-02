// vybe-test: csharp/csharp_foreach_iteration_order/foreach_iteration_order_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_foreach_iteration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// foreach_iteration_order
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
