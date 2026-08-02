// vybe-test: csharp/csharp_task_completion_matrix/task_completion_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_task_completion_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_completion_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[85] = 86; __Check((map.ContainsKey(85) && map[85] == 86).ToString(), "True");
