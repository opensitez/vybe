// vybe-test: csharp/csharp_task_timeout_matrix/task_timeout_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_task_timeout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_timeout_matrix
int seed = 86; int right = seed + 1; __Check((seed < right).ToString(), "True");
