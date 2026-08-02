// vybe-test: csharp/csharp_enum_metaprogramming/enum_parse_string_to_member
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Color{Red,Green,Blue} var c=(Color)System.Enum.Parse(typeof(Color),"Green"); __Check((c).ToString(), "Green");
