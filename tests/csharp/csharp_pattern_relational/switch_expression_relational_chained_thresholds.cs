// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_chained_thresholds
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int v=42; __Check((v switch{<10=>"xs",<100=>"md",_=>"lg"}).ToString(), "md");
