// vybe-test: csharp/csharp_enum_metaprogramming/enum_underlying_type_uint
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum UIntEnum:uint{Big=3000000000u} __Check((System.Enum.GetUnderlyingType(typeof(UIntEnum)).Name).ToString(), "UInt32");
