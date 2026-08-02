// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// for_loop_bounds
var map = new System.Collections.Generic.Dictionary<int, int>(); map[45] = 46; __Check((map.ContainsKey(45) && map[45] == 46).ToString(), "True");
