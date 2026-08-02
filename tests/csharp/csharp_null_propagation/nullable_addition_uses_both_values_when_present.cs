// vybe-test: csharp/csharp_null_propagation/nullable_addition_uses_both_values_when_present
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? left = 2; int? right = 5; __Check(((left ?? 0) + (right ?? 0)).ToString(), "7");
