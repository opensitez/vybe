// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_interface_implementation
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace App;
interface IRun { string Go(); }
class Runner : IRun { public string Go() => "go"; }
IRun r = new Runner();
__Check((r.Go()).ToString(), "go");
