// vybe-test: csharp/csharp_event_handler_matrix/event_handler_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_event_handler_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// event_handler_matrix
string feature = "event_handler_matrix:77"; __Check((feature.Length >= 1).ToString(), "True");
