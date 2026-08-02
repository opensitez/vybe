// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_lambda_in_method
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Lambda;
class Fn { public int Run() { System.Func<int, int> f = x => x + 1; return f(3); } }
__Check((new Fn().Run()).ToString(), "4");
