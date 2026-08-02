// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_static_readonly_field
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Config;
class App { public static readonly string Env = "prod"; }
__Check((App.Env).ToString(), "prod");
