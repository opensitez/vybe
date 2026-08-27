// vybe-test: csharp/csharp_local_function_static/static_local_function_min_of_two
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int Min(int a,int b){static int Pick(int x,int y)=>x<y?x:y; return Pick(a,b);}
__P((Min(3,9)).ToString());
__Check("3");

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
