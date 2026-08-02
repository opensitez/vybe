// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// record_shape_basics
var map = new System.Collections.Generic.Dictionary<int, int>(); map[39] = 40; __Check((map.ContainsKey(39) && map[39] == 40).ToString(), "True");
