// vybe-test: csharp/csharp_foreach_iteration_order/foreach_iteration_order_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_foreach_iteration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// foreach_iteration_order
var map = new System.Collections.Generic.Dictionary<int, int>(); map[46] = 47; __Check((map.ContainsKey(46) && map[46] == 47).ToString(), "True");
