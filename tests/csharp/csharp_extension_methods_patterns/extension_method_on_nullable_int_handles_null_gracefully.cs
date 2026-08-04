// vybe-test: csharp/csharp_extension_methods_patterns/extension_method_on_nullable_int_handles_null_gracefully
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods_patterns.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

static class NullableExt { public static int OrZero(this int? n) => n ?? 0; }
int? present=5, absent=null;
__P((present.OrZero()).ToString()); __P((absent.OrZero()).ToString());
__Check("5\n0");
