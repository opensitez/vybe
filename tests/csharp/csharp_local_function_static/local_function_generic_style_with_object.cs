// vybe-test: csharp/csharp_local_function_static/local_function_generic_style_with_object
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

string Format(int n){string F(int x)=>"n="+x; return F(n);}
__P((Format(42)).ToString());
__Check("n=42");

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
