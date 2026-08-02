// vybe-test: csharp/csharp_enum_metaprogramming/enum_is_defined_for_numeric_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Num{A=5,B=6} __Check((System.Enum.IsDefined(typeof(Num),5)).ToString(), "True");
