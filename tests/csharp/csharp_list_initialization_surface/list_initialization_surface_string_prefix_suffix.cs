// vybe-test: csharp/csharp_list_initialization_surface/list_initialization_surface_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_list_initialization_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_initialization_surface
string feature = "list_initialization_surface"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
