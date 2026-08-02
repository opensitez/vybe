// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_static_class_method
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Tools;
static class MathEx { public static int Double(int n) => n * 2; }
__Check((MathEx.Double(6)).ToString(), "12");
