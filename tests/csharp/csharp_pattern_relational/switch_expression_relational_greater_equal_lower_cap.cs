// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_greater_equal_lower_cap
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=1; __Check((n switch{>=1=>"ok",_=>"low"}).ToString(), "ok");
