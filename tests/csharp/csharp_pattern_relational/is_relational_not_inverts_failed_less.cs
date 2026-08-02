// vybe-test: csharp/csharp_pattern_relational/is_relational_not_inverts_failed_less
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=8; __Check((n is not <5).ToString(), "True");
