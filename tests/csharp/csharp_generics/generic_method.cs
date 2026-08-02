// vybe-test: csharp/csharp_generics/generic_method
// origin: languages/csharp/tests/csharp/test_csharp_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Utils {
    public static T Identity<T>(T value) { return value; }
}
__Check((Utils.Identity<int>(42)).ToString(), "42");
__Check((Utils.Identity<string>("hello")).ToString(), "hello");
