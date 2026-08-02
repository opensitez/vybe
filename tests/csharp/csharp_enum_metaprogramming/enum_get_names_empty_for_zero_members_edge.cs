// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_names_empty_for_zero_members_edge
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Edge{A} __Check((System.Enum.GetNames(typeof(Edge)).Length==1).ToString(), "True");
