// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_multiple_types_same_file
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Duo;
class A { public int Value = 1; }
class B { public int Value = 2; }
__Check((new A().Value).ToString(), "1"); __Check((new B().Value).ToString(), "2");
