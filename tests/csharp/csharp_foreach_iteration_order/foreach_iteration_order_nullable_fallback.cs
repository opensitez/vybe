// vybe-test: csharp/csharp_foreach_iteration_order/foreach_iteration_order_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_foreach_iteration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// foreach_iteration_order
int? maybe = null; int fallback = maybe ?? 46; __Check((fallback == 46).ToString(), "True");
