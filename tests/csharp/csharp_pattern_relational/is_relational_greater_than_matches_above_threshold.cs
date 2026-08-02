// vybe-test: csharp/csharp_pattern_relational/is_relational_greater_than_matches_above_threshold
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=5; __Check((n is >3).ToString(), "True");
