// vybe-test: csharp/csharp_pattern_property/is_property_pattern_float_field_literal
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Rate { public float R; } object o=new Rate{R=1.5f}; __Check((o is Rate{R:1.5f}).ToString(), "True");
