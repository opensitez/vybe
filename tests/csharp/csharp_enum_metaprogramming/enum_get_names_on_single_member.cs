// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_names_on_single_member
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Solo{Only} __Check((System.Enum.GetNames(typeof(Solo))[0]).ToString(), "Only");
