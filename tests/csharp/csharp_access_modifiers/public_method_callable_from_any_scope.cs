// vybe-test: csharp/csharp_access_modifiers/public_method_callable_from_any_scope
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

using static __Harness;

__P((new Service().Name()).ToString());
__Check("svc");

class Service{public string Name()=>"svc";}

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
