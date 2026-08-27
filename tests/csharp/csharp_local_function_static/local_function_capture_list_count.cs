// vybe-test: csharp/csharp_local_function_static/local_function_capture_list_count
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

var items=new System.Collections.Generic.List<int>{1,2,3}
;
int SizePlus(int n){int S(int x)=>items.Count+x; return S(n);}
__P((SizePlus(1)).ToString());
__Check("4");

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
