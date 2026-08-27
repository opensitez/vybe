// vybe-test: csharp/csharp_exceptions_hierarchy/custom_exception_with_inner_exception_chain
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

using static __Harness;

string r="";
try{throw new Outer(new System.ArgumentNullException("arg"));}
catch(Outer ex){r=ex.InnerException?.GetType().Name;}
__P((r).ToString());
__Check("ArgumentNullException");

class Outer:System.Exception{public Outer(System.Exception inner):base("outer",inner){}}

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
