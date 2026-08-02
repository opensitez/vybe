// vybe-test: csharp/csharp_type_conversions/boxing_nullable_with_value_prints_number
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? value = 13; object boxed = value; __Check((boxed).ToString(), "13");
