// vybe-test: csharp/csharp_pattern_property/is_property_pattern_zero_field_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Zero { public int Z; } object o=new Zero{Z=0}; __Check((o is Zero{Z:0}).ToString(), "True");
