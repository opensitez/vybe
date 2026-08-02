// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_predicate_apis
var map = new System.Collections.Generic.Dictionary<int, int>(); map[23] = 24; __Check((map.ContainsKey(23) && map[23] == 24).ToString(), "True");
