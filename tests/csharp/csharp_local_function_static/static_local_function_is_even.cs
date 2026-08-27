// vybe-test: csharp/csharp_local_function_static/static_local_function_is_even
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

bool Even(int n){static bool Check(int x)=>x%2==0; return Check(n);}
__P((Even(6)).ToString());
__Check("True");

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
