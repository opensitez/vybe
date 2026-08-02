// vybe-test: csharp/csharp_enum_metaprogramming/enum_flags_or_combine_numeric
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags] enum F{A=1,B=2,C=4} var v=F.A|F.C; __Check(((int)v).ToString(), "5");
