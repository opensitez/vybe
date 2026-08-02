// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_init_property_initializer
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Init;
class Config { public int Port { get; init; } = 80; }
__Check((new Config { Port = 443 }.Port).ToString(), "443");
