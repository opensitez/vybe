// vybe-test: csharp/csharp_explicit_typing_surface/explicit_typing_surface_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_explicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// explicit_typing_surface
string feature = "explicit_typing_surface"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
