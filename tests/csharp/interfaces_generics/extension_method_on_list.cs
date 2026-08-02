// vybe-test: csharp/interfaces_generics/extension_method_on_list
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class ListExtensions {
    public static string Join<T>(this List<T> list, string sep) {
        return string.Join(sep, list);
    }
}
var nums = new List<int> { 1, 2, 3, 4, 5 };
__Check((nums.Join(", ")).ToString(), "1, 2, 3, 4, 5");
