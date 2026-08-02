// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// monitor_lock_matrix
int seed = 84; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
