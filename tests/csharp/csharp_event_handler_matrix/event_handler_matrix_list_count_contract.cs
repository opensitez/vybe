// vybe-test: csharp/csharp_event_handler_matrix/event_handler_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_event_handler_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// event_handler_matrix
var values = new System.Collections.Generic.List<int> { 77, 78, 77 }; __Check((values.Count == 3).ToString(), "True");
