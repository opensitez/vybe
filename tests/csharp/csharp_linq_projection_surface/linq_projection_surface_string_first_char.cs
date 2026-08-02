// vybe-test: csharp/csharp_linq_projection_surface/linq_projection_surface_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_linq_projection_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_projection_surface
string feature = "linq_projection_surface"; __Check((feature[0] == feature[0]).ToString(), "True");
