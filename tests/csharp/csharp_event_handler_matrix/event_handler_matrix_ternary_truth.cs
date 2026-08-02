// vybe-test: csharp/csharp_event_handler_matrix/event_handler_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_event_handler_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// event_handler_matrix
int seed = 77; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
