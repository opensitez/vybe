// vybe-test: csharp/csharp_generics_advanced/default_of_generic_t_is_null_for_reference_types
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Null<T>() where T : class => default(T);
__Check((Null<string>() == null).ToString(), "True");
