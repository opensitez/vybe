// vybe-test: csharp/csharp_local_function_static/local_function_void_side_effect
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int Run(){int acc=0; void Bump(int n){acc+=n;} Bump(2); Bump(3); return acc;}
__P((Run()).ToString());
__Check("5");

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
