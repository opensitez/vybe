// vybe-test: csharp/csharp_generic_methods/generic_method_with_explicit_type_argument
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Box<T>(T v)=>v;
__Check((Box<int>(5)).ToString(), "5");
