// vybe-test: csharp/csharp_tuple_patterns/tuple_swap_without_temp_variable
// origin: languages/csharp/tests/csharp/test_csharp_tuple_patterns.rs

using static __Harness;

int x=1,y=2;
(x,y)=(y,x);
__P((x).ToString());
__P((y).ToString());
__Check("2\n1");

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
