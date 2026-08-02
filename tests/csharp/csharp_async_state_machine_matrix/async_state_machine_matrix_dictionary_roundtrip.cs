// vybe-test: csharp/csharp_async_state_machine_matrix/async_state_machine_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_async_state_machine_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_state_machine_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[88] = 89; __Check((map.ContainsKey(88) && map[88] == 89).ToString(), "True");
