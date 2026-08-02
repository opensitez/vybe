// vybe-test: csharp/csharp_extension_methods_patterns/extension_method_on_nullable_int_handles_null_gracefully
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class NullableExt { public static int OrZero(this int? n) => n ?? 0; }
int? present=5, absent=null;
__Check((present.OrZero()).ToString(), "5"); __Check((absent.OrZero()).ToString(), "0");
