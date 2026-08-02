// vybe-test: csharp/csharp_attribute_visibility_matrix/attribute_visibility_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_attribute_visibility_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// attribute_visibility_matrix
var tuple = (left: 93, right: 94); __Check((tuple.left < tuple.right).ToString(), "True");
