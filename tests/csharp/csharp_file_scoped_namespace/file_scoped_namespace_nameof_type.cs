// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_nameof_type
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Names;
class Widget { }
__Check((nameof(Widget)).ToString(), "Widget");
