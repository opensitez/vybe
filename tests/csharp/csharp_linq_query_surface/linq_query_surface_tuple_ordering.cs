// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_query_surface
var tuple = (left: 117, right: 118); __Check((tuple.left < tuple.right).ToString(), "True");
