// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_nested_class_access
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Core;
class Outer { public class Inner { public int Value = 7; } }
__Check((new Outer.Inner().Value).ToString(), "7");
