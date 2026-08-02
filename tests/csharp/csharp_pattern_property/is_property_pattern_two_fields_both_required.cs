// vybe-test: csharp/csharp_pattern_property/is_property_pattern_two_fields_both_required
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pair { public int A; public int B; } object o=new Pair{A=2,B=3}; __Check((o is Pair{A:2,B:3}).ToString(), "True");
