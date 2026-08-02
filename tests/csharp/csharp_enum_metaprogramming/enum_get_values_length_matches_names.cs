// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_values_length_matches_names
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Trio{X,Y,Z} __Check((System.Enum.GetValues(typeof(Trio)).Length).ToString(), "3");
