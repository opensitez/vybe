// vybe-test: csharp/csharp_list_initialization_surface/list_initialization_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_list_initialization_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_initialization_surface
var values = new System.Collections.Generic.List<int> { 30, 31, 30 }; __Check((values.Count == 3).ToString(), "True");
