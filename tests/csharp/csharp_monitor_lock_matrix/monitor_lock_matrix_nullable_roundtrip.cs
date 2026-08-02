// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// monitor_lock_matrix
int? maybe = 84; __Check((maybe.HasValue && maybe.Value == 84).ToString(), "True");
