// vybe-test: csharp/csharp_nullable_semantics/arithmetic_on_nullable_where_one_is_null_yields_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? a=3, b=null; __Check(((a+b).HasValue).ToString(), "False");
