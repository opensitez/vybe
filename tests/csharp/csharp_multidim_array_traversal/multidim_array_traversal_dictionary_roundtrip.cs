// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// multidim_array_traversal
var map = new System.Collections.Generic.Dictionary<int, int>(); map[29] = 30; __Check((map.ContainsKey(29) && map[29] == 30).ToString(), "True");
