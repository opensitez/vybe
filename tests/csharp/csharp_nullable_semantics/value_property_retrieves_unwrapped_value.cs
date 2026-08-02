// vybe-test: csharp/csharp_nullable_semantics/value_property_retrieves_unwrapped_value
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n = 42; __Check((n.Value).ToString(), "42");
