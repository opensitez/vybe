// vybe-test: csharp/csharp_integer_arithmetic/division_on_int64_framework_name_truncates
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Int64 a = 17, b = 5; __Check((a / b).ToString(), "3");
