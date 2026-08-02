// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// monitor_lock_matrix
int? maybe = null; int fallback = maybe ?? 84; __Check((fallback == 84).ToString(), "True");
