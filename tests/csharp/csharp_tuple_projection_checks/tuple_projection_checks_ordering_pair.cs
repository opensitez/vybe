// vybe-test: csharp/csharp_tuple_projection_checks/tuple_projection_checks_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_tuple_projection_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_projection_checks
int seed = 36; int right = seed + 1; __Check((seed < right).ToString(), "True");
