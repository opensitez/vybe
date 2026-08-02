// vybe-test: csharp/csharp_enum_metaprogramming/enum_to_string_member_name
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Letter{A,B,C} __Check((Letter.B.ToString()).ToString(), "B");
