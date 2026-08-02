// vybe-test: csharp/csharp_event_handler_matrix/event_handler_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_event_handler_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// event_handler_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[77] = 78; __Check((map.ContainsKey(77) && map[77] == 78).ToString(), "True");
