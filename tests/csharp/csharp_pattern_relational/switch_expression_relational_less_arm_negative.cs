// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_less_arm_negative
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=-4; __Check((n switch{<0=>"neg",0=>"zero",_=>"pos"}).ToString(), "neg");
