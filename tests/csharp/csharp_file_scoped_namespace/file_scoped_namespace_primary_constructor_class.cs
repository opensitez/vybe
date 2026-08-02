// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_primary_constructor_class
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Svc;
class Service(string name) { public string Name => name; }
__Check((new Service("api").Name).ToString(), "api");
