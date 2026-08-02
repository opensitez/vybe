// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

// short_circuit_logic_patterns
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
