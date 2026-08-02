// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_null_coalescing_assignment_skips_when_present
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n=3; n??=9; __Check((n).ToString(), "3");
