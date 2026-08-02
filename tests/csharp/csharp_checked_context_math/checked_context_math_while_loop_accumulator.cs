// vybe-test: csharp/csharp_checked_context_math/checked_context_math_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_checked_context_math.rs

// checked_context_math
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
