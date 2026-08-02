// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_names_preserves_declaration_order
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Order{Z,A,M} __Check((System.Enum.GetNames(typeof(Order))[1]).ToString(), "A");
