// vybe-test: csharp/csharp_enum_flags_operations/flags_enum_has_flag_detects_combined_permission
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags]
enum Perm { Read = 1, Write = 2, Execute = 4 }
var value = Perm.Read | Perm.Write;
__Check((value.HasFlag(Perm.Write)).ToString(), "True");
