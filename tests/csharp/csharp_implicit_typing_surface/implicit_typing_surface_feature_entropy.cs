// vybe-test: csharp/csharp_implicit_typing_surface/implicit_typing_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_implicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// implicit_typing_surface
string feature = "implicit_typing_surface:59"; __Check((feature.Length >= 1).ToString(), "True");
