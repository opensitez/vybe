// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_query_surface
var values = new System.Collections.Generic.List<int> { 117, 118, 117 }; __Check((values.Count == 3).ToString(), "True");
