// vybe-test: csharp/csharp_enum_metaprogramming/enum_parse_with_explicit_values
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Http{Ok=200,NotFound=404} var v=(Http)System.Enum.Parse(typeof(Http),"NotFound"); __Check(((int)v).ToString(), "404");
