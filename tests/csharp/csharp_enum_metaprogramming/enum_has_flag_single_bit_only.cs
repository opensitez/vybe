// vybe-test: csharp/csharp_enum_metaprogramming/enum_has_flag_single_bit_only
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags] enum Bit{One=1,Two=2} var v=Bit.Two; __Check((v.HasFlag(Bit.One)).ToString(), "False"); __Check((v.HasFlag(Bit.Two)).ToString(), "True");
