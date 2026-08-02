// vybe-test: csharp/csharp_enum_flags_operations/plain_enum_cast_to_int_preserves_underlying_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Level { Low = 1, Mid = 5, High = 9 }
__Check(((int)Level.Mid).ToString(), "5");
