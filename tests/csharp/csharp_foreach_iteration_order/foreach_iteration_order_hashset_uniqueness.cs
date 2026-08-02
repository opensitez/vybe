// vybe-test: csharp/csharp_foreach_iteration_order/foreach_iteration_order_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_foreach_iteration_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// foreach_iteration_order
var set = new System.Collections.Generic.HashSet<int>(); set.Add(46); set.Add(46); __Check((set.Count == 1).ToString(), "True");
