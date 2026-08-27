// vybe-test: csharp/csharp_with_expression_records/with_list_element
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var list=new System.Collections.Generic.List<V>{new V(1),new V(2)}
;
list[1]=list[1] with{N=9}
;
__P((list[1].N).ToString());
__Check("9");

record V(int N);

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
