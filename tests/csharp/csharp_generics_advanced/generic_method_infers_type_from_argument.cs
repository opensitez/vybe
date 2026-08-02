// vybe-test: csharp/csharp_generics_advanced/generic_method_infers_type_from_argument
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Identity<T>(T value) => value;
__Check((Identity(99)).ToString(), "99");
__Check((Identity("hi")).ToString(), "hi");
