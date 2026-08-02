// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

// partial_type_behavior
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
