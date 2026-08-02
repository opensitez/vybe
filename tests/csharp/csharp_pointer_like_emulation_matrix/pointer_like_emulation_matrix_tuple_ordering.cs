// vybe-test: csharp/csharp_pointer_like_emulation_matrix/pointer_like_emulation_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_pointer_like_emulation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pointer_like_emulation_matrix
var tuple = (left: 114, right: 115); __Check((tuple.left < tuple.right).ToString(), "True");
