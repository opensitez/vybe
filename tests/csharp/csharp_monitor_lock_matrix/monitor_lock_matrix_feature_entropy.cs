// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// monitor_lock_matrix
string feature = "monitor_lock_matrix:84"; __Check((feature.Length >= 1).ToString(), "True");
