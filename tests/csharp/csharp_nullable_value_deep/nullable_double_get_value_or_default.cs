// vybe-test: csharp/csharp_nullable_value_deep/nullable_double_get_value_or_default
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double? d=null; __Check((d.GetValueOrDefault(3.14)).ToString(), "3.14");
