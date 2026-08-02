// vybe-test: csharp/csharp_timespan_arithmetic_matrix/timespan_arithmetic_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// timespan_arithmetic_matrix
var tuple = (left: 95, right: 96); __Check((tuple.left < tuple.right).ToString(), "True");
