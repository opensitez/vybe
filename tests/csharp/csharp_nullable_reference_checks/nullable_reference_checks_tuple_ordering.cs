// vybe-test: csharp/csharp_nullable_reference_checks/nullable_reference_checks_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_nullable_reference_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_reference_checks
var tuple = (left: 58, right: 59); __Check((tuple.left < tuple.right).ToString(), "True");
