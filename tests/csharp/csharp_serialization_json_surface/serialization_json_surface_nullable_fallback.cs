// vybe-test: csharp/csharp_serialization_json_surface/serialization_json_surface_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_serialization_json_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// serialization_json_surface
int? maybe = null; int fallback = maybe ?? 91; __Check((fallback == 91).ToString(), "True");
