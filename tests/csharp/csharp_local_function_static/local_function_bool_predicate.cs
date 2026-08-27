// vybe-test: csharp/csharp_local_function_static/local_function_bool_predicate
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

bool AllPositive(int a,int b){bool Check(int x,int y)=>x>0&&y>0; return Check(a,b);}
__P((AllPositive(1,2)).ToString());
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
