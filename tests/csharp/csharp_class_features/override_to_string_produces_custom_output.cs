// vybe-test: csharp/csharp_class_features/override_to_string_produces_custom_output
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

using static __Harness;

__P((new Color(255,0,128)).ToString());
__Check("rgb(255,0,128)");

class Color{int R,G,B;public Color(int r,int g,int b){R=r;G=g;B=b;}
public override string ToString()=>$"rgb({R},{G},{B})";}

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
