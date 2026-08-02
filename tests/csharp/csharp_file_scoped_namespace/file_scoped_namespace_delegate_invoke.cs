// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_delegate_invoke
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Fn;
delegate int Getter();
Getter g = () => 42;
__Check((g()).ToString(), "42");
