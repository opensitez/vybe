// vybe-test: csharp/csharp_enum_metaprogramming/enum_compare_equal_members
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Eq{X,Y} __Check((Eq.X==Eq.X).ToString(), "True"); __Check((Eq.X==Eq.Y).ToString(), "False");
