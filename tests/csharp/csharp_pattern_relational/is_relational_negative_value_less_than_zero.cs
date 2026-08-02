// vybe-test: csharp/csharp_pattern_relational/is_relational_negative_value_less_than_zero
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=-1; __Check((n is <0).ToString(), "True");
