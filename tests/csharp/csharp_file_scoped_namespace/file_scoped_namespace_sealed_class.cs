// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_sealed_class
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Seal;
sealed class Token { public int Id = 1; }
__Check((new Token().Id).ToString(), "1");
