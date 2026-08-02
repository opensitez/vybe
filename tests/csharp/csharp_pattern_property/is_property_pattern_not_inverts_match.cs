// vybe-test: csharp/csharp_pattern_property/is_property_pattern_not_inverts_match
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { public int V; } object o=new Box{V=1}; __Check((o is not Box{V:2}).ToString(), "True");
