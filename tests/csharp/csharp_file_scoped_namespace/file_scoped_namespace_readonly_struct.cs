// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_readonly_struct
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Immut;
readonly struct Pair { public readonly int A; public readonly int B; public Pair(int a, int b) { A = a; B = b; } }
__Check((new Pair(2, 3).A + new Pair(2, 3).B).ToString(), "5");
