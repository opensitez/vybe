// vybe-test: csharp/csharp_task_completion_matrix/task_completion_matrix_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_task_completion_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_completion_matrix
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
