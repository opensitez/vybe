// vybe-test: csharp/csharp_interlocked_atomic_matrix/interlocked_atomic_matrix_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic_matrix.rs

// interlocked_atomic_matrix
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
