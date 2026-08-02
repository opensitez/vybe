// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_construction_matrix
var tuple = (left: 94, right: 95); __Check((tuple.left < tuple.right).ToString(), "True");
