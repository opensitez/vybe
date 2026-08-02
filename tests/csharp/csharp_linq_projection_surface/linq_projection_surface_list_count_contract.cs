// vybe-test: csharp/csharp_linq_projection_surface/linq_projection_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_linq_projection_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_projection_surface
var values = new System.Collections.Generic.List<int> { 118, 119, 118 }; __Check((values.Count == 3).ToString(), "True");
