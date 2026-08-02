// vybe-test: csharp/csharp_numeric_ops/integer_plus_double_widens_to_double
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int i=3; double d=1.5;
__Check((i+d).ToString(), "4.5");
