// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_collection_expression
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Coll;
int[] data = [1, 2, 3];
__Check((data[1]).ToString(), "2");
