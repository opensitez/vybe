// vybe-test: csharp/csharp_enum_metaprogramming/enum_cast_to_short_underlying
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum ShortEnum:short{Neg=-1,Pos=1} __Check(((short)ShortEnum.Neg).ToString(), "-1");
