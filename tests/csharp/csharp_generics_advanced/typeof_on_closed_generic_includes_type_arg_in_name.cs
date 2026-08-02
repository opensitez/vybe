// vybe-test: csharp/csharp_generics_advanced/typeof_on_closed_generic_includes_type_arg_in_name
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((typeof(System.Collections.Generic.List<int>).IsGenericType).ToString(), "True");
