// vybe-test: csharp/csharp_local_function_static/static_local_function_called_from_sibling_local
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int Pipeline(int n){static int Double(int x)=>x*2; int Wrap(int v)=>Double(v)+1; return Wrap(n);}
__P((Pipeline(5)).ToString());
__Check("11");

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
