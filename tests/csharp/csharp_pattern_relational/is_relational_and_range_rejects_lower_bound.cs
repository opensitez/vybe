// vybe-test: csharp/csharp_pattern_relational/is_relational_and_range_rejects_lower_bound
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=10; __Check((n is >10 and <20).ToString(), "False");
