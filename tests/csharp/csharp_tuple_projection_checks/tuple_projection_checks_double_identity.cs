// vybe-test: csharp/csharp_tuple_projection_checks/tuple_projection_checks_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_tuple_projection_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_projection_checks
double seed = 36; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
