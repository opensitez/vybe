// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_immutability_checks
int seed = 18; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
