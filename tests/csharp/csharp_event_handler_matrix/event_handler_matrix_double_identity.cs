// vybe-test: csharp/csharp_event_handler_matrix/event_handler_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_event_handler_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// event_handler_matrix
double seed = 77; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
