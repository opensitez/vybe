// vybe-test: csharp/csharp_pattern_property/is_property_pattern_long_field_literal
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Span { public long L; } object o=new Span{L=1000L}; __Check((o is Span{L:1000L}).ToString(), "True");
