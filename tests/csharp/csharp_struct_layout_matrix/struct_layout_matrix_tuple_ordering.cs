// vybe-test: csharp/csharp_struct_layout_matrix/struct_layout_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_struct_layout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// struct_layout_matrix
var tuple = (left: 113, right: 114); __Check((tuple.left < tuple.right).ToString(), "True");
