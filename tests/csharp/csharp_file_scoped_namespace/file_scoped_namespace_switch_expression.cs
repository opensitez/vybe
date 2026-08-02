// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_switch_expression
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Switch;
string Label(int n) => n switch { 1 => "one", _ => "other" };
__Check((Label(1)).ToString(), "one");
