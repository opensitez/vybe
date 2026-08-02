// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// short_circuit_logic_patterns
var set = new System.Collections.Generic.HashSet<int>(); set.Add(14); set.Add(14); __Check((set.Count == 1).ToString(), "True");
