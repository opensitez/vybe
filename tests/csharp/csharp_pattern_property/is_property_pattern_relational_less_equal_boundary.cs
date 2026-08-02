// vybe-test: csharp/csharp_pattern_property/is_property_pattern_relational_less_equal_boundary
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Level { public int Tier; } object o=new Level{Tier=3}; __Check((o is Level{Tier:<=3}).ToString(), "True");
