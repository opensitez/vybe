// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// short_circuit_logic_patterns
var values = new System.Collections.Generic.List<int> { 14, 15, 14 }; __Check((values.Count == 3).ToString(), "True");
