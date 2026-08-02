// vybe-test: csharp/csharp_pattern_property/is_property_pattern_and_two_relational_fields
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Range { public int Lo; public int Hi; } object o=new Range{Lo=5,Hi=15}; __Check((o is Range{Lo:>0,Hi:<20}).ToString(), "True");
