// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// partial_type_behavior
double seed = 70; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
