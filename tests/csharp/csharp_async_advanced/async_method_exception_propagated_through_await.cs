// vybe-test: csharp/csharp_async_advanced/async_method_exception_propagated_through_await
// origin: languages/csharp/tests/csharp/test_csharp_async_advanced.rs

using static __Harness;

async System.Threading.Tasks.Task Fail()=>throw new System.Exception("async fail");
string msg="";
try{await Fail();}
catch(System.Exception ex){msg=ex.Message;}
__P((msg).ToString());
__Check("async fail");

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
