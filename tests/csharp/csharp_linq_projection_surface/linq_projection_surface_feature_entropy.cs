// vybe-test: csharp/csharp_linq_projection_surface/linq_projection_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_linq_projection_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_projection_surface
string feature = "linq_projection_surface:118"; __Check((feature.Length >= 1).ToString(), "True");
