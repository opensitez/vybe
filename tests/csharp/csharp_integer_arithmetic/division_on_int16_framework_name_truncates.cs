// vybe-test: csharp/csharp_integer_arithmetic/division_on_int16_framework_name_truncates
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Int16 a = 9, b = 4; __Check((a / b).ToString(), "2");
