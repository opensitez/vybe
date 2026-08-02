// vybe-test: csharp/csharp_if_else_branching/if_else_branching_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
var map = new System.Collections.Generic.Dictionary<int, int>(); map[44] = 45; __Check((map.ContainsKey(44) && map[44] == 45).ToString(), "True");
