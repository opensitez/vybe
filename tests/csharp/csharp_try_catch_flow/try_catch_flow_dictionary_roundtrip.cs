// vybe-test: csharp/csharp_try_catch_flow/try_catch_flow_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_try_catch_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// try_catch_flow
var map = new System.Collections.Generic.Dictionary<int, int>(); map[51] = 52; __Check((map.ContainsKey(51) && map[51] == 52).ToString(), "True");
