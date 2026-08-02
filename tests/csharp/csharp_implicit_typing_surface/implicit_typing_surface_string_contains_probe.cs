// vybe-test: csharp/csharp_implicit_typing_surface/implicit_typing_surface_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_implicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// implicit_typing_surface
string feature = "implicit_typing_surface"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
