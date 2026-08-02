// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// partial_type_behavior
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
