// vybe-test: csharp/csharp_interlocked_atomic_matrix/interlocked_atomic_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interlocked_atomic_matrix
var values = new System.Collections.Generic.List<int> { 83, 84, 83 }; __Check((values.Count == 3).ToString(), "True");
