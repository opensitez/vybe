// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_list_generic
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Lists;
var list = new System.Collections.Generic.List<int> { 1, 2 };
__Check((list.Count).ToString(), "2");
