// vybe-test: csharp/csharp_pattern_is_checks/pattern_is_checks_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_pattern_is_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_is_checks
var tuple = (left: 41, right: 42); __Check((tuple.left < tuple.right).ToString(), "True");
