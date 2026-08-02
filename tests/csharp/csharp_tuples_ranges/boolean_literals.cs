// vybe-test: csharp/csharp_tuples_ranges/boolean_literals
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((true).ToString(), "True");
__Check((false).ToString(), "False");
__Check((true && false).ToString(), "False");
__Check((true || false).ToString(), "True");
