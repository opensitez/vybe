// vybe-test: csharp/csharp_pattern_relational/is_relational_on_long_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

long x=5000L; __Check((x is >=1000L).ToString(), "True");
