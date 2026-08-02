// vybe-test: csharp/csharp_serialization_json_surface/serialization_json_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_serialization_json_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// serialization_json_surface
var values = new System.Collections.Generic.List<int> { 91, 92, 91 }; __Check((values.Count == 3).ToString(), "True");
