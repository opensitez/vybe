// vybe-test: csharp/csharp_list_initialization_surface/list_initialization_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_list_initialization_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_initialization_surface
string feature = "list_initialization_surface:30"; __Check((feature.Length >= 1).ToString(), "True");
