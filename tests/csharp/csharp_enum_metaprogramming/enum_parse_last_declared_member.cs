// vybe-test: csharp/csharp_enum_metaprogramming/enum_parse_last_declared_member
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Level{Low,Mid,High} var v=(Level)System.Enum.Parse(typeof(Level),"High"); __Check((v).ToString(), "High");
