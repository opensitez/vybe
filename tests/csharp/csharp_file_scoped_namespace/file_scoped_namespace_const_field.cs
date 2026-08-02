// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_const_field
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Consts;
class Limits { public const int Max = 100; }
__Check((Limits.Max).ToString(), "100");
