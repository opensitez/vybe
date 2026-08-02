// vybe-test: csharp/csharp_task_timeout_matrix/task_timeout_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_task_timeout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_timeout_matrix
var values = new System.Collections.Generic.List<int> { 86, 87, 86 }; __Check((values.Count == 3).ToString(), "True");
