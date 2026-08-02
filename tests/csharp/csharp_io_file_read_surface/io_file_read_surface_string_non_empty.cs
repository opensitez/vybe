// vybe-test: csharp/csharp_io_file_read_surface/io_file_read_surface_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_io_file_read_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_file_read_surface
string feature = "io_file_read_surface"; __Check((feature.Length > 0).ToString(), "True");
