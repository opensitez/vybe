// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_class_method_call
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Demo;
class Worker { public string Run() => "ok"; }
__Check((new Worker().Run()).ToString(), "ok");
