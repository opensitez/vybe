// vybe-test: csharp/csharp_guid_parse_matrix/guid_parse_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// guid_parse_matrix
var tuple = (left: 97, right: 98); __Check((tuple.left < tuple.right).ToString(), "True");
