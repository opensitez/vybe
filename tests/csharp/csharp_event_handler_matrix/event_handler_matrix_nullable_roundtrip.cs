// vybe-test: csharp/csharp_event_handler_matrix/event_handler_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_event_handler_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// event_handler_matrix
int? maybe = 77; __Check((maybe.HasValue && maybe.Value == 77).ToString(), "True");
