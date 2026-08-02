// vybe-test: csharp/csharp_boolean_algebra_surfaces/boolean_algebra_surfaces_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_boolean_algebra_surfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boolean_algebra_surfaces
var values = new System.Collections.Generic.List<int> { 11, 12, 11 }; __Check((values.Count == 3).ToString(), "True");
