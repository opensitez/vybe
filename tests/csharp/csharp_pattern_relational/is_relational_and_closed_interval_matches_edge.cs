// vybe-test: csharp/csharp_pattern_relational/is_relational_and_closed_interval_matches_edge
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=80; __Check((n is >=80 and <=89).ToString(), "True");
