// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_dotted_namespace_name
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Acme.Widgets;
class Widget { public string Name => "w"; }
__Check((new Widget().Name).ToString(), "w");
