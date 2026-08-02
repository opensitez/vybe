// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_enum_underlying_int
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Tier { Low=1, Mid=5, High=10 } var t=Tier.Mid; __Check((t switch{>=Tier.Mid=>"up",_=>"down"}).ToString(), "up");
