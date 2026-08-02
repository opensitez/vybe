// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// short_circuit_logic_patterns
string feature = "short_circuit_logic_patterns"; __Check((feature[0] == feature[0]).ToString(), "True");
