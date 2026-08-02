// vybe-test: csharp/csharp_enum_operations/flags_enum_has_flag_detects_combined_bit
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags] enum Perm{None=0,Read=1,Write=2,Execute=4}
var p = Perm.Read | Perm.Write;
__Check((p.HasFlag(Perm.Read)).ToString(), "True");
__Check((p.HasFlag(Perm.Execute)).ToString(), "False");
