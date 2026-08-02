// vybe-test: csharp/csharp_pattern_relational/is_relational_and_three_part_window
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=50; __Check((n is >=40 and <=60 and !=55).ToString(), "True");
