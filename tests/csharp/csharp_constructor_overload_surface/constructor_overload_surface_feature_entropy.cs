// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_overload_surface
string feature = "constructor_overload_surface:67"; __Check((feature.Length >= 1).ToString(), "True");
