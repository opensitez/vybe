// vybe-test: csharp/csharp_floating_point_literals_surface/floating_point_literals_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_floating_point_literals_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// floating_point_literals_surface
var values = new System.Collections.Generic.List<int> { 16, 17, 16 }; __Check((values.Count == 3).ToString(), "True");
