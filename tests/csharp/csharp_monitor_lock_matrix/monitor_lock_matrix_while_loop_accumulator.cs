// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

// monitor_lock_matrix
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
