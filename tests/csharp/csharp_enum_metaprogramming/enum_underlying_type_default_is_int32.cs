// vybe-test: csharp/csharp_enum_metaprogramming/enum_underlying_type_default_is_int32
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Plain{One} __Check((System.Enum.GetUnderlyingType(typeof(Plain)).Name).ToString(), "Int32");
