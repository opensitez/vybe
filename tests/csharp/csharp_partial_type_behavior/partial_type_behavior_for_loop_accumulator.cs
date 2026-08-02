// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

// partial_type_behavior
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
