// vybe-test: csharp/csharp_event_handler_matrix/event_handler_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_event_handler_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// event_handler_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(77); set.Add(77); __Check((set.Count == 1).ToString(), "True");
