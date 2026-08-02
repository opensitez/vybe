// vybe-test: csharp/csharp_floating_point_literals_surface/floating_point_literals_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_floating_point_literals_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// floating_point_literals_surface
string feature = "floating_point_literals_surface:16"; __Check((feature.Length >= 1).ToString(), "True");
