// vybe-test: csharp/csharp_generics_advanced/default_of_generic_t_is_zero_for_value_types
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Zero<T>() => default(T);
__Check((Zero<int>()).ToString(), "0");
__Check((Zero<bool>()).ToString(), "False");
