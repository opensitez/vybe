// vybe-test: csharp/csharp_linq_projection_surface/linq_projection_surface_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_linq_projection_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_projection_surface
var tuple = (left: 118, right: 119); __Check((tuple.left < tuple.right).ToString(), "True");
