// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_enum_member
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Flags;
enum Mode { Off, On }
__Check((Mode.On).ToString(), "On");
