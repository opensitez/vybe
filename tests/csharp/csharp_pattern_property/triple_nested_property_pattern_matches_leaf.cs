// vybe-test: csharp/csharp_pattern_property/triple_nested_property_pattern_matches_leaf
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Leaf { public int V; } class Mid { public Leaf L; } class Root { public Mid M; } object o=new Root{M=new Mid{L=new Leaf{V=4}}}; if(o is Root{M:{L:{V:4}}}) __P(("deep").ToString());
__Check("deep");
