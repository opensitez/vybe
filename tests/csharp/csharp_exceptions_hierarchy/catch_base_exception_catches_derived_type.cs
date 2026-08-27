// vybe-test: csharp/csharp_exceptions_hierarchy/catch_base_exception_catches_derived_type
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

using static __Harness;

string r="";
try{int[] a=new int[3]; var _=a[10];}
catch(System.Exception ex){r=ex.GetType().Name;}
__P((r).ToString());
__Check("IndexOutOfRangeException");

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
