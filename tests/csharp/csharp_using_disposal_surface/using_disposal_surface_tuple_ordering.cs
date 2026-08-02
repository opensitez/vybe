// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// using_disposal_surface
var tuple = (left: 52, right: 53); __Check((tuple.left < tuple.right).ToString(), "True");
