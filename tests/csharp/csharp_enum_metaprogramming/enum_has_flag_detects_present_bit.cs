// vybe-test: csharp/csharp_enum_metaprogramming/enum_has_flag_detects_present_bit
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags] enum Perm{Read=1,Write=2} var p=Perm.Read|Perm.Write; __Check((p.HasFlag(Perm.Read)).ToString(), "True");
