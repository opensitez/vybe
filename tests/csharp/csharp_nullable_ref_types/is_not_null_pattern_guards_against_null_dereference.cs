// vybe-test: csharp/csharp_nullable_ref_types/is_not_null_pattern_guards_against_null_dereference
// origin: languages/csharp/tests/csharp/test_csharp_nullable_ref_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string? s="hello";
if(s is not null) __Check((s.Length).ToString(), "5");
