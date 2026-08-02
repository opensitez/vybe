// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// short_circuit_logic_patterns
var tuple = (left: 14, right: 15); __Check((tuple.left < tuple.right).ToString(), "True");
