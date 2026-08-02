// vybe-test: csharp/csharp_conversion_methods/convert_to_double_from_int
// origin: languages/csharp/tests/csharp/test_csharp_conversion_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double d=System.Convert.ToDouble(7);
__Check((d).ToString(), "7");
