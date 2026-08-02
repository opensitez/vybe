// vybe-test: csharp/csharp_linq_projection_surface/linq_projection_surface_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_linq_projection_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_projection_surface
string feature = "linq_projection_surface"; __Check((feature.Length > 0).ToString(), "True");
