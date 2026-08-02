// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// monitor_lock_matrix
var tuple = (left: 84, right: 85); __Check((tuple.left < tuple.right).ToString(), "True");
