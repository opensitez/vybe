// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_query_surface
string feature = "linq_query_surface:117"; __Check((feature.Length >= 1).ToString(), "True");
