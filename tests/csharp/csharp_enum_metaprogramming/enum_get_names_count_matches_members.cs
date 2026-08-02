// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_names_count_matches_members
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Size{Small,Medium,Large,Extra} __Check((System.Enum.GetNames(typeof(Size)).Length).ToString(), "4");
