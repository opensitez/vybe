// vybe-test: csharp/csharp_enum_metaprogramming/enum_flags_and_mask
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags] enum F{A=1,B=2,C=4} var v=(F.A|F.B|F.C)&F.B; __Check(((int)v).ToString(), "2");
