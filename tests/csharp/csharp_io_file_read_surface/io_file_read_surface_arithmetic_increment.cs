// vybe-test: csharp/csharp_io_file_read_surface/io_file_read_surface_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_io_file_read_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_file_read_surface
int seed = 89; __Check((seed + 1 > seed).ToString(), "True");
