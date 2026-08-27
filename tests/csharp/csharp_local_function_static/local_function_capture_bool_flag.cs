// vybe-test: csharp/csharp_local_function_static/local_function_capture_bool_flag
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

bool enabled=true;
int Gate(int n){int G(int x)=>enabled?x:-x; return G(n);}
__P((Gate(7)).ToString());
__Check("7");

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
