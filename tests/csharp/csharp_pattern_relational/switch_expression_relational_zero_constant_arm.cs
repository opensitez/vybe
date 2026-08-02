// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_zero_constant_arm
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=0; __Check((n switch{<0=>"neg",0=>"zero",>0=>"pos"}).ToString(), "zero");
