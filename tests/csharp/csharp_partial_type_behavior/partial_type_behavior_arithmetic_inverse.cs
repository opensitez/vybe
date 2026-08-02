// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// partial_type_behavior
int seed = 70; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
