// vybe-test: csharp/csharp_pattern_property/is_property_pattern_string_literal_mismatch
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Tag { public string Name; } object o=new Tag{Name="alpha"}; __Check((o is Tag{Name:"beta"}).ToString(), "False");
