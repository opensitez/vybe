// vybe-test: csharp/csharp_local_function_static/local_function_capture_enum
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

Mode state=Mode.On;
int Code(int n){int C(int x)=>state==Mode.On?x:0; return C(n);}
__P((Code(5)).ToString());
__Check("5");

enum Mode{On,Off}

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
