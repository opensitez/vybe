// vybe-test: csharp/csharp_event_handler_matrix/event_handler_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_event_handler_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// event_handler_matrix
int seed = 77; int right = seed + 1; __Check((seed < right).ToString(), "True");
