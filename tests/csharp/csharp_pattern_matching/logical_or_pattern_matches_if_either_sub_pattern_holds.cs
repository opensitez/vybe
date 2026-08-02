// vybe-test: csharp/csharp_pattern_matching/logical_or_pattern_matches_if_either_sub_pattern_holds
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 5;
__Check((n is 3 or 5 or 7).ToString(), "True");
