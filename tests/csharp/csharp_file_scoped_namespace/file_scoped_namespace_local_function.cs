// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_local_function
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Local;
int Twice(int n) { int Double(int x) => x * 2; return Double(n); }
__Check((Twice(5)).ToString(), "10");
