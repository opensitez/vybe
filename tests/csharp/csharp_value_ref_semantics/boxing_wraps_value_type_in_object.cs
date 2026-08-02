// vybe-test: csharp/csharp_value_ref_semantics/boxing_wraps_value_type_in_object
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=42; object o=n;
__Check((o).ToString(), "42"); __Check((o is int).ToString(), "True");
