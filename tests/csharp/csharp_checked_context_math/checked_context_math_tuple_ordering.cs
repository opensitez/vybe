// vybe-test: csharp/csharp_checked_context_math/checked_context_math_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_checked_context_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// checked_context_math
var tuple = (left: 12, right: 13); __Check((tuple.left < tuple.right).ToString(), "True");
