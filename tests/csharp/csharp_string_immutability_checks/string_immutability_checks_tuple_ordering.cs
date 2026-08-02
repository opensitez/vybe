// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_immutability_checks
var tuple = (left: 18, right: 19); __Check((tuple.left < tuple.right).ToString(), "True");
