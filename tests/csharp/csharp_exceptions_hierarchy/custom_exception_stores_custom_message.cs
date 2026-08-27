// vybe-test: csharp/csharp_exceptions_hierarchy/custom_exception_stores_custom_message
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

using static __Harness;

string r="";
try{throw new AppEx("fail");}
catch(AppEx ex){r=ex.Message;}
__P((r).ToString());
__Check("fail");

class AppEx:System.Exception{public AppEx(string m):base(m){}}

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
