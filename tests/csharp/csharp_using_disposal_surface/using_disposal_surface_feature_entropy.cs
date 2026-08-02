// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// using_disposal_surface
string feature = "using_disposal_surface:52"; __Check((feature.Length >= 1).ToString(), "True");
