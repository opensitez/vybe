// vybe-test: csharp/csharp_enum_flags_operations/flags_enum_none_is_zero_and_or_identity
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags]
enum Perm { None = 0, Read = 1 }
var value = Perm.None | Perm.Read;
__Check(((int)value).ToString(), "1");
