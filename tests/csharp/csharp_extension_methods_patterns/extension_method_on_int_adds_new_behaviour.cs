// vybe-test: csharp/csharp_extension_methods_patterns/extension_method_on_int_adds_new_behaviour
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class IntExt { public static bool IsEven(this int n) => n%2==0; }
__Check((4.IsEven()).ToString(), "True"); __Check((3.IsEven()).ToString(), "False");
