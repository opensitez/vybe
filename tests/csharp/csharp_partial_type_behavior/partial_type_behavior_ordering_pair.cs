// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// partial_type_behavior
int seed = 70; int right = seed + 1; __Check((seed < right).ToString(), "True");
