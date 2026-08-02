// vybe-test: csharp/csharp_nullable_ref_types/nullable_reference_getvalueorlength_via_null_coalescing
// origin: languages/csharp/tests/csharp/test_csharp_nullable_ref_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string? s=null;
int len=s?.Length??-1;
__Check((len).ToString(), "-1");
