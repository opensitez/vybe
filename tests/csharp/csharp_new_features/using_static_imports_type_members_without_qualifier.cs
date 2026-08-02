// vybe-test: csharp/csharp_new_features/using_static_imports_type_members_without_qualifier
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using static System.Math;
__Check((Sqrt(16)).ToString(), "4");
