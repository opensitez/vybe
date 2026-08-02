// vybe-test: csharp/csharp_nullable_semantics/comparing_two_nulls_returns_true_for_equality
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? a=null, b=null; __Check((a==b).ToString(), "True");
