// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

// threading_pool_matrix
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
