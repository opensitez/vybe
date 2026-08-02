// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_query_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[117] = 118; __Check((map.ContainsKey(117) && map[117] == 118).ToString(), "True");
