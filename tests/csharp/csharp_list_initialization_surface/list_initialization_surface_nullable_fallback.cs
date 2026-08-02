// vybe-test: csharp/csharp_list_initialization_surface/list_initialization_surface_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_list_initialization_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_initialization_surface
int? maybe = null; int fallback = maybe ?? 30; __Check((fallback == 30).ToString(), "True");
