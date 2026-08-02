// vybe-test: csharp/csharp_pattern_property/is_property_pattern_negative_field_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Delta { public int D; } object o=new Delta{D=-5}; __Check((o is Delta{D:-5}).ToString(), "True");
