// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_public_class_from_same_file
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Pub;
public class Visible { public string Text => "seen"; }
__Check((new Visible().Text).ToString(), "seen");
