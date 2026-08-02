// vybe-test: csharp/csharp_numeric_precision/big_integer_can_hold_arbitrarily_large_values
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var n=System.Numerics.BigInteger.Pow(10,30);
__Check((n.ToString().StartsWith("1")).ToString(), "True");
