// vybe-test: csharp/csharp_type_conversions/convert_class_creates_double_from_integer
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Convert.ToDouble(5)).ToString(), "5");
