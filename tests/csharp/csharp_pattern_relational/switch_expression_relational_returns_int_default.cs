// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_returns_int_default
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=3; __Check((n switch{>10=>20,>5=>10,_=>0}).ToString(), "0");
