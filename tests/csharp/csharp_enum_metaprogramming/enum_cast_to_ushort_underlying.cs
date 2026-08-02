// vybe-test: csharp/csharp_enum_metaprogramming/enum_cast_to_ushort_underlying
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum UShortEnum:ushort{Max=65535} __Check(((ushort)UShortEnum.Max).ToString(), "65535");
