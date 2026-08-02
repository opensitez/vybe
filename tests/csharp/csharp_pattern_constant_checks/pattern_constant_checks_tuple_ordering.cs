// vybe-test: csharp/csharp_pattern_constant_checks/pattern_constant_checks_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_pattern_constant_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_constant_checks
var tuple = (left: 40, right: 41); __Check((tuple.left < tuple.right).ToString(), "True");
