// vybe-test: csharp/csharp_pattern_property/triple_nested_property_pattern_matches_leaf
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Leaf { public int V; } class Mid { public Leaf L; } class Root { public Mid M; } object o=new Root{M=new Mid{L=new Leaf{V=4}}}; if(o is Root{M:{L:{V:4}}}) __Check(("deep").ToString(), "deep");
