// vybe-test: csharp/csharp_io_file_read_surface/io_file_read_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_io_file_read_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_file_read_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(89); set.Add(89); __Check((set.Count == 1).ToString(), "True");
