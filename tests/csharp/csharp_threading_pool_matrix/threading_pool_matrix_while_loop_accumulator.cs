// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

// threading_pool_matrix
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
