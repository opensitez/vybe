// vybe-test: csharp/csharp_delegates_advanced/delegate_stored_in_dictionary_and_dispatched_by_key
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

using static __Harness;

var ops=new System.Collections.Generic.Dictionary<string,System.Func<int,int,int>>{
    {"add",(a,b)=>a+b},
    {"mul",(a,b)=>a*b}
}
;
__P((ops["add"](3,4)).ToString());
__P((ops["mul"](3,4)).ToString());
__Check("7\n12");

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
