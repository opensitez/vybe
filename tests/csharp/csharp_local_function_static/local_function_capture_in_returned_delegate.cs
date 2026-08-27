// vybe-test: csharp/csharp_local_function_static/local_function_capture_in_returned_delegate
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

System.Func<int,int> MakeScaler(int factor){int Scale(int x)=>x*factor; return Scale;}
__P((MakeScaler(4)(6)).ToString());
__Check("24");

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
