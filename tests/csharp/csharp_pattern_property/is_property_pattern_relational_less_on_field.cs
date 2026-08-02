// vybe-test: csharp/csharp_pattern_property/is_property_pattern_relational_less_on_field
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Score { public int Points; } object o=new Score{Points=40}; __Check((o is Score{Points:<50}).ToString(), "True");
