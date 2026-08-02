// vybe-test: csharp/csharp_pattern_relational/is_relational_less_than_rejects_equal_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=5; __Check((n is <5).ToString(), "False");
