// vybe-test: csharp/csharp_io_file_read_surface/io_file_read_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_io_file_read_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_file_read_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[89] = 90; __Check((map.ContainsKey(89) && map[89] == 90).ToString(), "True");
