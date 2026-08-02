// vybe-test: csharp/csharp_pattern_property/is_property_pattern_partial_single_field_ignores_rest
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Wide { public int A; public int B; public int C; } object o=new Wide{A=1,B=2,C=3}; __Check((o is Wide{A:1}).ToString(), "True");
