// vybe-test: csharp/csharp_numeric_precision/float_is_32_bit_and_less_precise_than_double
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

float f=1.0f/3.0f;
double d=1.0/3.0;
__Check((f==(float)d).ToString(), "True");
