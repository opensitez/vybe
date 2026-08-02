// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// monitor_lock_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[84] = 85; __Check((map.ContainsKey(84) && map[84] == 85).ToString(), "True");
