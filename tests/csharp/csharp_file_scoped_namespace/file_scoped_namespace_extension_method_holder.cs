// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_extension_method_holder
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Ext;
static class IntExt { public static int Inc(this int n) => n + 1; }
__Check((4.Inc()).ToString(), "5");
