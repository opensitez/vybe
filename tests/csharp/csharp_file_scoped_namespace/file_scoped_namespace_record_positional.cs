// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_record_positional
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Data;
record Pair(int A, int B);
__Check((new Pair(1, 2).A).ToString(), "1");
