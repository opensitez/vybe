// vybe-test: csharp/csharp_event_handler_matrix/event_handler_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_event_handler_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// event_handler_matrix
int? maybe = null; int fallback = maybe ?? 77; __Check((fallback == 77).ToString(), "True");
