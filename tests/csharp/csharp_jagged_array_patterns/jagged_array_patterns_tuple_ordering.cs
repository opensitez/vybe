// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
var tuple = (left: 28, right: 29); __Check((tuple.left < tuple.right).ToString(), "True");
