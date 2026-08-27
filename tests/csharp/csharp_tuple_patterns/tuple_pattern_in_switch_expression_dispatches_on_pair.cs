// vybe-test: csharp/csharp_tuple_patterns/tuple_pattern_in_switch_expression_dispatches_on_pair
// origin: languages/csharp/tests/csharp/test_csharp_tuple_patterns.rs

using static __Harness;

string Classify(int x,int y)=>(x,y) switch{
    (0,0)=>"origin",
    (>0,0)=>"pos-x",
    (0,>0)=>"pos-y",
    _=>"other"}
;
__P((Classify(0,0)).ToString());
__P((Classify(3,0)).ToString());
__P((Classify(0,5)).ToString());
__P((Classify(1,1)).ToString());
__Check("origin\npos-x\npos-y\nother");

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
