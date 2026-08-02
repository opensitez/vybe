// vybe-test: csharp/csharp_namespace_aliases/namespace_can_contain_static_helper_class
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Demo.Tools { public static class MathEx { public static int Double(int value) { return value * 2; } } } __Check((Demo.Tools.MathEx.Double(6)).ToString(), "12");
