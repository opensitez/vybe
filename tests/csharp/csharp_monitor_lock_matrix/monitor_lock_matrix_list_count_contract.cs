// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// monitor_lock_matrix
var values = new System.Collections.Generic.List<int> { 84, 85, 84 }; __Check((values.Count == 3).ToString(), "True");
