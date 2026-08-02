// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_property_getter
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Props;
class Box { public int Size { get; } = 10; }
__Check((new Box().Size).ToString(), "10");
