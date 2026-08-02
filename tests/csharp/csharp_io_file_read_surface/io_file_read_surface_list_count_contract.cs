// vybe-test: csharp/csharp_io_file_read_surface/io_file_read_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_io_file_read_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_file_read_surface
var values = new System.Collections.Generic.List<int> { 89, 90, 89 }; __Check((values.Count == 3).ToString(), "True");
