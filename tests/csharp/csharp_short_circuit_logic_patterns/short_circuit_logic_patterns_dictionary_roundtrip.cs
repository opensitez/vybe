// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// short_circuit_logic_patterns
var map = new System.Collections.Generic.Dictionary<int, int>(); map[14] = 15; __Check((map.ContainsKey(14) && map[14] == 15).ToString(), "True");
