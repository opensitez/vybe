// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// partial_type_behavior
int? maybe = 70; __Check((maybe.HasValue && maybe.Value == 70).ToString(), "True");
