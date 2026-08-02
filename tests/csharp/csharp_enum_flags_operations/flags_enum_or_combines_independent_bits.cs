// vybe-test: csharp/csharp_enum_flags_operations/flags_enum_or_combines_independent_bits
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags]
enum Perm { None = 0, Read = 1, Write = 2 }
var value = Perm.Read | Perm.Write;
__Check(((int)value).ToString(), "3");
