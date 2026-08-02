// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_nested_selector
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int a=2,b=3; __Check(((a+b) switch{>4=>"big",_=>"small"}).ToString(), "big");
