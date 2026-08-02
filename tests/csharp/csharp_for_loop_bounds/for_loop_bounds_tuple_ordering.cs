// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// for_loop_bounds
var tuple = (left: 45, right: 46); __Check((tuple.left < tuple.right).ToString(), "True");
