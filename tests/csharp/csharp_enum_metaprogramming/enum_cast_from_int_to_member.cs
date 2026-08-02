// vybe-test: csharp/csharp_enum_metaprogramming/enum_cast_from_int_to_member
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Code{A=1,B=2} var c=(Code)2; __Check((c).ToString(), "B");
