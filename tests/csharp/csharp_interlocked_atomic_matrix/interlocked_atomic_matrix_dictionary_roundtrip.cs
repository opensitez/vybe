// vybe-test: csharp/csharp_interlocked_atomic_matrix/interlocked_atomic_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interlocked_atomic_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[83] = 84; __Check((map.ContainsKey(83) && map[83] == 84).ToString(), "True");
