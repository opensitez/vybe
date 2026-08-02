// vybe-test: csharp/csharp_type_conversions/convert_class_parses_integer_from_string
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Convert.ToInt32("42") + 8).ToString(), "50");
