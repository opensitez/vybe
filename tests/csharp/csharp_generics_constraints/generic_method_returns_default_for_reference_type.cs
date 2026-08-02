// vybe-test: csharp/csharp_generics_constraints/generic_method_returns_default_for_reference_type
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Empty<T>() where T : class { return default(T); } __Check((Empty<string>() is null).ToString(), "True");
