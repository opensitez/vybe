// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_less_equal_upper_cap
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=100; __Check((n switch{<=100=>"ok",_=>"high"}).ToString(), "ok");
