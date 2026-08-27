// vybe-test: csharp/csharp_local_function_static/local_function_capture_nullable_int_has_value
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int? opt=7;
int Bump(int n){int B(int x)=>x+(opt??0); return B(n);}
__P((Bump(1)).ToString());
__Check("8");

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
