// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_deep_namespace_path
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace A.B.C.D;
class Node { public string Name => "deep"; }
__Check((new Node().Name).ToString(), "deep");
