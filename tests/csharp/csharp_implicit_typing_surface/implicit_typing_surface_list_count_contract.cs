// vybe-test: csharp/csharp_implicit_typing_surface/implicit_typing_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_implicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// implicit_typing_surface
var values = new System.Collections.Generic.List<int> { 59, 60, 59 }; __Check((values.Count == 3).ToString(), "True");
