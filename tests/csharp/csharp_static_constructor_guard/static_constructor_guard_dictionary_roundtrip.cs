// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
var map = new System.Collections.Generic.Dictionary<int, int>(); map[69] = 70; __Check((map.ContainsKey(69) && map[69] == 70).ToString(), "True");
