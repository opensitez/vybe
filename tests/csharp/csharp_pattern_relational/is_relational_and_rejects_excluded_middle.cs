// vybe-test: csharp/csharp_pattern_relational/is_relational_and_rejects_excluded_middle
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=55; __Check((n is >=40 and <=60 and !=55).ToString(), "False");
