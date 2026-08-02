// vybe-test: csharp/csharp_pattern_matching/logical_and_pattern_requires_both_sub_patterns
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 15;
__Check((n is > 10 and < 20).ToString(), "True");
