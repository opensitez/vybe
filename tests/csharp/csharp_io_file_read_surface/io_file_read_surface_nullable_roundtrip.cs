// vybe-test: csharp/csharp_io_file_read_surface/io_file_read_surface_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_io_file_read_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_file_read_surface
int? maybe = 89; __Check((maybe.HasValue && maybe.Value == 89).ToString(), "True");
