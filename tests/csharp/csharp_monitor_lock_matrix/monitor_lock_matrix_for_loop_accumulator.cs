// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

// monitor_lock_matrix
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
