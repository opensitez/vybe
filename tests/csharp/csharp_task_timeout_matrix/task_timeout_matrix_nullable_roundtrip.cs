// vybe-test: csharp/csharp_task_timeout_matrix/task_timeout_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_task_timeout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_timeout_matrix
int? maybe = 86; __Check((maybe.HasValue && maybe.Value == 86).ToString(), "True");
