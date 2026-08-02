// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_generic_class
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Gen;
class Box<T> { public T Item; }
var b = new Box<int> { Item = 9 };
__Check((b.Item).ToString(), "9");
