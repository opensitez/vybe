// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_overload_surface
var values = new System.Collections.Generic.List<int> { 67, 68, 67 }; __Check((values.Count == 3).ToString(), "True");
