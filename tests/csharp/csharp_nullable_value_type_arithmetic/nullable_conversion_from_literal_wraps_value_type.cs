// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_conversion_from_literal_wraps_value_type
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? boxed = 42;
__Check((boxed is int).ToString(), "True");
__Check(((int)boxed).ToString(), "42");
