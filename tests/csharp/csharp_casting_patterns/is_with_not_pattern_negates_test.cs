// vybe-test: csharp/csharp_casting_patterns/is_with_not_pattern_negates_test
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o="hello";
__Check((o is not int).ToString(), "True");
