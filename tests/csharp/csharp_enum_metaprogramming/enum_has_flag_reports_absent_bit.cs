// vybe-test: csharp/csharp_enum_metaprogramming/enum_has_flag_reports_absent_bit
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags] enum Perm{Read=1,Write=2,Execute=4} var p=Perm.Read; __Check((p.HasFlag(Perm.Execute)).ToString(), "False");
