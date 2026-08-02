// vybe-test: csharp/csharp_nullable_ref_types/nullable_string_can_hold_null_value
// origin: languages/csharp/tests/csharp/test_csharp_nullable_ref_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string? s=null;
__Check((s==null).ToString(), "True");
