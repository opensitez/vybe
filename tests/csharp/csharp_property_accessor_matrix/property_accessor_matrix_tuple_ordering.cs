// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// property_accessor_matrix
var tuple = (left: 64, right: 65); __Check((tuple.left < tuple.right).ToString(), "True");
