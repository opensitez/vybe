// vybe-test: csharp/csharp_switch_type_patterns/is_not_pattern_negates_type_test
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object value = 3.14;
__Check((value is not int).ToString(), "True");
