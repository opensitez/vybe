// vybe-test: csharp/csharp_loops/continue_skips_rest_of_body_and_re_evaluates_condition
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

using static __Harness;

int s=0;
for(int i=0;i<5;i++) { if(i%2==0) continue; s+=i; }
__P((s).ToString());
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
