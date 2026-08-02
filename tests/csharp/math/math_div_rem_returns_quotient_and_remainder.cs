// vybe-test: csharp/math/math_div_rem_returns_quotient_and_remainder
// origin: languages/csharp/tests/csharp/test_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int remainder;
var quotient = System.Math.DivRem(17, 5, out remainder);
__Check((quotient).ToString(), "3");
__Check((remainder).ToString(), "2");
