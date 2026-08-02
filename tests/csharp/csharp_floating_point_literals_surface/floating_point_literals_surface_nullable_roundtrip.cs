// vybe-test: csharp/csharp_floating_point_literals_surface/floating_point_literals_surface_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_floating_point_literals_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// floating_point_literals_surface
int? maybe = 16; __Check((maybe.HasValue && maybe.Value == 16).ToString(), "True");
