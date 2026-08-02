// vybe-test: csharp/csharp_task_timeout_matrix/task_timeout_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_task_timeout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_timeout_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[86] = 87; __Check((map.ContainsKey(86) && map[86] == 87).ToString(), "True");
