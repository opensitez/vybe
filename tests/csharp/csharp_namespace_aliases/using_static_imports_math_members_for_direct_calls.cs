// vybe-test: csharp/csharp_namespace_aliases/using_static_imports_math_members_for_direct_calls
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using static System.Math; __Check((Max(3, 9)).ToString(), "9");
