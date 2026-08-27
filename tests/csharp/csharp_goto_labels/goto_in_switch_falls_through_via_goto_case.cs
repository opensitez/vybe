// vybe-test: csharp/csharp_goto_labels/goto_in_switch_falls_through_via_goto_case
// origin: languages/csharp/tests/csharp/test_csharp_goto_labels.rs

using static __Harness;

int n=1;
string r="";
switch(n){
    case 1: r+="one"; goto case 2;
    case 2: r+="two"; break;
    case 3: r+="three"; break;
}
__P((r).ToString());
__Check("onetwo");

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
