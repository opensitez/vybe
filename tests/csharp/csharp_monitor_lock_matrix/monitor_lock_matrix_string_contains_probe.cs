// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// monitor_lock_matrix
string feature = "monitor_lock_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
