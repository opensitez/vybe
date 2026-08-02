// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_null_coalescing_right_when_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n=null; __Check((n??42).ToString(), "42");
