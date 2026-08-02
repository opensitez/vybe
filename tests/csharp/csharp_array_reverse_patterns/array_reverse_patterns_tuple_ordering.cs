// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_reverse_patterns
var tuple = (left: 27, right: 28); __Check((tuple.left < tuple.right).ToString(), "True");
