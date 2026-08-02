// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_required_property
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Req;
class User { public required string Name { get; set; } }
__Check((new User { Name = "Ada" }.Name).ToString(), "Ada");
