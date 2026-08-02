// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_method_with_params
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Calc;
class Adder { public int Sum(int a, int b) => a + b; }
__Check((new Adder().Sum(2, 3)).ToString(), "5");
