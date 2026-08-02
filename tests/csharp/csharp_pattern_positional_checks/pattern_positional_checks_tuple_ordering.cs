// vybe-test: csharp/csharp_pattern_positional_checks/pattern_positional_checks_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_pattern_positional_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_positional_checks
var tuple = (left: 115, right: 116); __Check((tuple.left < tuple.right).ToString(), "True");
