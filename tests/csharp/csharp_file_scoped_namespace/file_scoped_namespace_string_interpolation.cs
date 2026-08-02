// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_string_interpolation
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Text;
class Tag { public string Label(string name) => $"hi {name}"; }
__Check((new Tag().Label("Ann")).ToString(), "hi Ann");
