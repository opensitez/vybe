// vybe-test: csharp/csharp_string_split_join/string_split_join_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
var map = new System.Collections.Generic.Dictionary<int, int>(); map[21] = 22; __Check((map.ContainsKey(21) && map[21] == 22).ToString(), "True");
