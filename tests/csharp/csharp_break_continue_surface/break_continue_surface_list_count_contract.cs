// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// break_continue_surface
var values = new System.Collections.Generic.List<int> { 49, 50, 49 }; __Check((values.Count == 3).ToString(), "True");
