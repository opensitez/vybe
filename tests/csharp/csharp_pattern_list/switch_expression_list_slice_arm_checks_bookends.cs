// vybe-test: csharp/csharp_pattern_list/switch_expression_list_slice_arm_checks_bookends
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

string Edge(int[] a)=>a switch{[1,..,9]=>"book",_=>"plain"}
;
__P((Edge(new[]{1,5,9})).ToString());
__Check("book");

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
