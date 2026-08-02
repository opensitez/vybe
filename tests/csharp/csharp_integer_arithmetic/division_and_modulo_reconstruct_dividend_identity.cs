// vybe-test: csharp/csharp_integer_arithmetic/division_and_modulo_reconstruct_dividend_identity
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int dividend = 17; int divisor = 5; __Check((dividend / divisor * divisor + dividend % divisor).ToString(), "17");
