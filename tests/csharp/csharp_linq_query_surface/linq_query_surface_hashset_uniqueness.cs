// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_query_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(117); set.Add(117); __Check((set.Count == 1).ToString(), "True");
