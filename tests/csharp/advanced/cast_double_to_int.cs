// vybe-test: csharp/advanced/cast_double_to_int
// origin: languages/csharp/tests/csharp/test_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double d = 3.14;
        __Check((d).ToString(), "3.14");
