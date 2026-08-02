// vybe-test: csharp/csharp_pattern_relational/is_relational_on_double_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double d=3.14; __Check((d is >3.0).ToString(), "True");
