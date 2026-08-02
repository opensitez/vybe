// vybe-test: csharp/csharp_io_file_read_surface/io_file_read_surface_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_io_file_read_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_file_read_surface
var tuple = (left: 89, right: 90); __Check((tuple.left < tuple.right).ToString(), "True");
