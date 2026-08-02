// vybe-test: csharp/csharp_io_file_read_surface/io_file_read_surface_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_io_file_read_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_file_read_surface
int seed = 89; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
