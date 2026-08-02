// vybe-test: csharp/csharp_nullable_semantics/arithmetic_on_two_nullable_ints_with_values_produces_nullable_result
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? a=3, b=4; int? c=a+b; __Check((c).ToString(), "7");
