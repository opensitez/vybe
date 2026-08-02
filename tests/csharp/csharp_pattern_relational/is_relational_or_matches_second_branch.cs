// vybe-test: csharp/csharp_pattern_relational/is_relational_or_matches_second_branch
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=15; __Check((n is <0 or >10).ToString(), "True");
