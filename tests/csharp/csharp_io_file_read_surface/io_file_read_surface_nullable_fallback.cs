// vybe-test: csharp/csharp_io_file_read_surface/io_file_read_surface_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_io_file_read_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_file_read_surface
int? maybe = null; int fallback = maybe ?? 89; __Check((fallback == 89).ToString(), "True");
