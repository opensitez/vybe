// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_when_guard_fallback
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=2; __Check((n switch{int x when x>5 and x<10=>"mid",_=>"other"}).ToString(), "other");
