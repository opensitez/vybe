// vybe-test: csharp/csharp_array_copy_behavior/array_copy_behavior_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_array_copy_behavior.rs

// array_copy_behavior
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
