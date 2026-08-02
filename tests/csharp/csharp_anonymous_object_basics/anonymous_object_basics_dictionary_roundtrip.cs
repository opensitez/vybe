// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// anonymous_object_basics
var map = new System.Collections.Generic.Dictionary<int, int>(); map[38] = 39; __Check((map.ContainsKey(38) && map[38] == 39).ToString(), "True");
