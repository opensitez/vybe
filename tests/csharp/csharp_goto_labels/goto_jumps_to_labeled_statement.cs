// vybe-test: csharp/csharp_goto_labels/goto_jumps_to_labeled_statement
// origin: languages/csharp/tests/csharp/test_csharp_goto_labels.rs

using static __Harness;

int i=0;
start:
if(i<5){i++;goto start;}
__P((i).ToString());
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
