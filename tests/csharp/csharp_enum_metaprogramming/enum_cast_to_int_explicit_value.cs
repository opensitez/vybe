// vybe-test: csharp/csharp_enum_metaprogramming/enum_cast_to_int_explicit_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Code{A=10,B=20} __Check(((int)Code.B).ToString(), "20");
