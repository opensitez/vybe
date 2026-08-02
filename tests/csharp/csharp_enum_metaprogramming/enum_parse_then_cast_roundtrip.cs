// vybe-test: csharp/csharp_enum_metaprogramming/enum_parse_then_cast_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Round{A=11,B=22} var p=(Round)System.Enum.Parse(typeof(Round),"B"); __Check(((int)p).ToString(), "22");
