// vybe-test: csharp/csharp_pattern_list/switch_statement_list_pattern_case_with_var_capture
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

int[] data=new[]{3,9}
;
string tag="";
switch(data){case[var a,var b]:tag=(a+b).ToString();break;default:tag="0";break;}
__P((tag).ToString());
__Check("12");

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
