// vybe-test: csharp/csharp_pattern_property/triple_nested_property_pattern_matches_leaf
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object o=new Root{M=new Mid{L=new Leaf{V=4}}}
;
if(o is Root{M:{L:{V:4}}}) __P(("deep").ToString());
__Check("deep");

class Leaf { public int V; }

class Mid { public Leaf L; }

class Root { public Mid M; }

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
