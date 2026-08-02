// vybe-test: csharp/csharp_explicit_typing_surface/explicit_typing_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_explicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// explicit_typing_surface
var values = new System.Collections.Generic.List<int> { 60, 61, 60 }; __Check((values.Count == 3).ToString(), "True");
