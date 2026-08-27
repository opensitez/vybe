// vybe-test: csharp/csharp_local_function_static/local_function_returns_local_delegate
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

System.Func<int,int> MakeAdder(int n){int Add(int x)=>x+n; return Add;}
var add5=MakeAdder(5);
__P((add5(10)).ToString());
__Check("15");

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
