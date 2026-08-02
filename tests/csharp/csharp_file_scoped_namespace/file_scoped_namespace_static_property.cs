// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_static_property
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Static;
class Cache { public static int Size { get; set; } = 5; }
__Check((Cache.Size).ToString(), "5");
