// vybe-test: csharp/csharp_nested_type_access/nested_access_private_nested_enum_via_public_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Status().Read()).ToString());
__Check("0");

class Status{enum Code{Ok=0,Fail=1} public int Read()=>(int)Code.Ok;}

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
