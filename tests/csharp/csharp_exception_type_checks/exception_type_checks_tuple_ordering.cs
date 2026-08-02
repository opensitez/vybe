// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// exception_type_checks
var tuple = (left: 53, right: 54); __Check((tuple.left < tuple.right).ToString(), "True");
