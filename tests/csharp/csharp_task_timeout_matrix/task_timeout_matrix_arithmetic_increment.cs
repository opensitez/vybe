// vybe-test: csharp/csharp_task_timeout_matrix/task_timeout_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_task_timeout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_timeout_matrix
int seed = 86; __Check((seed + 1 > seed).ToString(), "True");
