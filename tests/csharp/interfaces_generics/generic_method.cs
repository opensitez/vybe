// vybe-test: csharp/interfaces_generics/generic_method
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Utils {
    public static T Max<T>(T a, T b) where T : IComparable<T> {
        return a.CompareTo(b) > 0 ? a : b;
    }
}
__Check((Utils.Max(3, 7)).ToString(), "7");
__Check((Utils.Max("apple", "banana")).ToString(), "banana");
