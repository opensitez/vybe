// vybe-test: csharp/csharp_enum_metaprogramming/enum_cast_to_byte_underlying
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Tiny:byte{A=200,B=201} __Check(((byte)Tiny.B).ToString(), "201");
