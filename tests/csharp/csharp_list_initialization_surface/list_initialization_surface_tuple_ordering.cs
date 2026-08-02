// vybe-test: csharp/csharp_list_initialization_surface/list_initialization_surface_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_list_initialization_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_initialization_surface
var tuple = (left: 30, right: 31); __Check((tuple.left < tuple.right).ToString(), "True");
