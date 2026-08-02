// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// monitor_lock_matrix
double seed = 84; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
