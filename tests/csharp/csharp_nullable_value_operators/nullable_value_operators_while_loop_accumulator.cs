// vybe-test: csharp/csharp_nullable_value_operators/nullable_value_operators_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_operators.rs

// nullable_value_operators
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
