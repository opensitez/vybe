// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// break_continue_surface
var tuple = (left: 49, right: 50); __Check((tuple.left < tuple.right).ToString(), "True");
