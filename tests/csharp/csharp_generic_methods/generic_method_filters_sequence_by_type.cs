// vybe-test: csharp/csharp_generic_methods/generic_method_filters_sequence_by_type
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

using static __Harness;

System.Collections.Generic.IEnumerable<T> FilterType<T>(object[] items){
    foreach(var i in items) if(i is T t) yield return t;
}
var items=new object[]{1,"a",2,"b",3}
;
int count=0;
foreach(var s in FilterType<string>(items)) count++;
__P((count).ToString());
__Check("2");

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
