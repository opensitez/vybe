// vybe-test: csharp/csharp_enum_operations/enum_explicit_value_tostring_probe
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Num{X=7} __Check((Num.X.ToString()).ToString(), "X");
