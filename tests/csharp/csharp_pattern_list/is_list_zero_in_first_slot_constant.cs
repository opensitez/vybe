// vybe-test: csharp/csharp_pattern_list/is_list_zero_in_first_slot_constant
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

int[] data=new[]{0,42}
;
if(data is [0,var v]) __P((v).ToString());
__Check("42");

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
