// vybe-test: csharp/csharp_goto_labels/continue_skips_rest_of_current_iteration
// origin: languages/csharp/tests/csharp/test_csharp_goto_labels.rs

using static __Harness;

int sum=0;
for(int i=1;i<=10;i++){
    if(i%2==0) continue;
    sum+=i;
}
__P((sum).ToString());
__Check("25");

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
