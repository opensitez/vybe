// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// using_disposal_surface
string feature = "using_disposal_surface"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
