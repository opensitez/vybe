// vybe-test: csharp/csharp_tuple_projection_checks/tuple_projection_checks_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_tuple_projection_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_projection_checks
var tuple = (left: 36, right: 37); __Check((tuple.left < tuple.right).ToString(), "True");
