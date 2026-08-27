// vybe-test: csharp/csharp_local_function_static/local_function_tail_call_style_sum
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int Sum(int n){int Loop(int i,int acc)=>i>n?acc:Loop(i+1,acc+i); return Loop(1,0);}
__P((Sum(4)).ToString());
__Check("10");

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
