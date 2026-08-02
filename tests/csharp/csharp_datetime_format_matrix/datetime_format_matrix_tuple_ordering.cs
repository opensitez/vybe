// vybe-test: csharp/csharp_datetime_format_matrix/datetime_format_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_datetime_format_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_format_matrix
var tuple = (left: 96, right: 97); __Check((tuple.left < tuple.right).ToString(), "True");
