// vybe-test: csharp/csharp_pattern_property/is_property_pattern_double_field_literal
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Measure { public double M; } object o=new Measure{M=2.5}; __Check((o is Measure{M:2.5}).ToString(), "True");
