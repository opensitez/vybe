// vybe-test: csharp/csharp_operators/is_not_pattern_negates_type_test
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object value = 7;
__Check((value is not string).ToString(), "True");
