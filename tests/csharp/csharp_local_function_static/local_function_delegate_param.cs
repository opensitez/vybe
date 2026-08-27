// vybe-test: csharp/csharp_local_function_static/local_function_delegate_param
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int Apply(int n,System.Func<int,int> op){int Wrap(int x)=>op(x)+1; return Wrap(n);}
__P((Apply(4,x=>x*2)).ToString());
__Check("9");

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
